"""
lib_parser.py
Smart Spice .lib 파일 파서

문법 참조:
  .LIB <libname>
    .MODEL <name> <type>
    + param1=val1 param2=val2 ...
    .PARAM var1=v1 var2=v2
  .ENDL <libname>
  *  주석 (줄 첫 글자가 *)
  $ 인라인 주석
"""
import re
from collections import OrderedDict
from typing import List, Tuple

from data_model import LibFile, LibBlock, ModelEntry, ParamEntry, DirectiveEntry

# 인라인 주석 제거 ($ 이후 제거), 문자열 리터럴 내부 보호
_INLINE_COMMENT_RE = re.compile(r'\$.*$')


def _strip_inline_comment(line: str) -> str:
    """$ 이후의 인라인 주석을 제거합니다 (문자열 외부에서만)."""
    return _INLINE_COMMENT_RE.sub('', line).rstrip()


def _parse_param_pairs(text: str) -> tuple[OrderedDict, bool, bool]:
    """
    'param1=val1 param2={val2+1} ...' 형식의 단일 텍스트 라인을
    OrderedDict({'param1': 'val1', 'param2': '{val2+1}'}) 객체로 파싱합니다.
    
    이때 파라미터 리스트 전체가 괄호로 둘러싸여 있는 경우(예: '(VTH0=0.4 TOX=1e-8)')
    괄호를 벗겨내고, 시작/종료 괄호가 있었는지에 대한 boolean 플래그를 함께 반환하여
    나중에 저장할 때 괄호 덩어리를 다시 복원할 수 있게 합니다.
    """
    params = OrderedDict()
    text = text.strip()
    
    open_paren = False
    close_paren = False
    
    if text.startswith('('):
        open_paren = True
        text = text[1:].strip()
        
    if text.endswith(')'):
        close_paren = True
        text = text[:-1].strip()

    if not text:
        return params, open_paren, close_paren

    # 토큰 분리 정규식: '파라미터명=값' 쌍을 추출합니다. (값이 중괄호 {} 로 묶인 수식을 포함할 수 있음)
    # 예: vth0=0.45 tox={tox_var*1.1} level=3
    pattern = re.compile(
        r'(\w+)\s*=\s*'           # 파라미터 이름(알파베틱 문자와 숫자 조합)과 등호(=)
        r'('
        r'\{[^}]*\}'              # 그룹 1: 중괄호로 덮인 수학 수식 (예: {var + 1})
        r'|'
        r'[^\s={}]+'              # 그룹 2: 일반 숫자 문자열 (공백이나 = 기호가 없는 토큰)
        r')'
    )
    for m in pattern.finditer(text):
        name = m.group(1).strip()
        value = m.group(2).strip()
        
        # 만약 값 중에 닫는 괄호가 섞여 들어왔다면 제거하고 close_paren 표시
        if value.endswith(')'):
            value = value[:-1]
            close_paren = True
            
        params[name] = value

    return params, open_paren, close_paren


def _join_continuation_lines(raw_lines: List[str]) -> List[str]:
    """
    + 로 시작하는 continuation line을 이전 줄에 합칩니다.
    주석 줄(* 로 시작)은 그대로 유지합니다.
    """
    joined: List[str] = []
    for raw in raw_lines:
        stripped = raw.strip()
        if not stripped:
            joined.append('')
            continue
        if stripped.startswith('+'):
            # continuation: 이전 실질 내용 줄에 이어 붙임
            rest = stripped[1:].strip()
            # 이전 줄을 찾아서 이어붙이기
            if joined:
                joined[-1] = joined[-1].rstrip() + ' ' + rest
            else:
                joined.append(rest)
        else:
            joined.append(stripped)
    return joined


def parse_lib(filepath: str) -> LibFile:
    """
    .lib 파일을 파싱하여 LibFile 객체를 반환합니다.
    """
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        raw_lines = f.readlines()

    lib_file = LibFile(filepath=filepath)

    # 1단계: continuation line 합치기
    lines = _join_continuation_lines(raw_lines)

    # 2단계: 순차 파싱
    current_lib: LibBlock = None
    current_model: ModelEntry = None
    pending_comments: List[str] = []
    in_lib_block = False

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # 빈 줄
        if not stripped:
            i += 1
            continue

        # 주석 줄 (* 시작)
        if stripped.startswith('*'):
            pending_comments.append(stripped)
            i += 1
            continue

        # 인라인 주석 제거
        stripped = _strip_inline_comment(stripped)
        if not stripped:
            i += 1
            continue

        upper = stripped.upper()

        # [조건 1] .LIB 시작 지시어 처리
        # .LIBRARY 는 일반 라이브러리 참조 구문일 수 있으므로 제외합니다.
        if upper.startswith('.LIB') and not upper.startswith('.LIBRARY'):
            parts = stripped.split(None, 1)
            lib_name = parts[1].strip() if len(parts) > 1 else 'UNNAMED'
            
            # 새 LIB 블록 객체 생성
            current_lib = LibBlock(name=lib_name, leading_comments=pending_comments)
            pending_comments = []
            current_model = None
            in_lib_block = True
            i += 1
            continue

        # [조건 2] .ENDL (라이브러리 끝) 처리
        if upper.startswith('.ENDL'):
            if current_lib is not None:
                # 현재 기록 중인 라이브러리 블록을 전체 파일 구조(lib_blocks)에 추가
                lib_file.lib_blocks.append(current_lib)
                
            # 상태 초기화
            current_lib = None
            current_model = None
            in_lib_block = False
            i += 1
            continue

        # [조건 3] .MODEL 처리
        if upper.startswith('.MODEL'):
            parts = stripped.split(None, 3)
            # 기본 형식: .MODEL <name> <type> [params...]
            model_name = parts[1] if len(parts) > 1 else 'UNKNOWN'
            model_type = parts[2] if len(parts) > 2 else ''
            param_text = parts[3] if len(parts) > 3 else ''
            
            # 모델 타입에 '(' 가 붙어있을 수 있으므로 이를 분리합니다. (예: .MODEL NMOS_1 (NMOS ...)
            m_open_paren = False
            if '(' in model_type:
                model_type = model_type.replace('(', '').strip()
                m_open_paren = True
            
            # 파라미터 텍스트 문자열에 처음 시작하는 괄호 '(' 분리
            if param_text.startswith('('):
                param_text = param_text[1:].strip()
                m_open_paren = True
                
            # 내부 함수 호출로 키워드 매칭 및 후행 괄호 파싱
            model_params, p_open, p_close = _parse_param_pairs(param_text)
            
            current_model = ModelEntry(
                name=model_name,
                model_type=model_type,
                params=model_params,
                comment_lines=pending_comments,
                open_paren=m_open_paren or p_open,
                close_paren=p_close
            )
            pending_comments = []
            if current_lib is not None:
                current_lib.models.append(current_model)
            i += 1
            continue

        # [조건 4] .PARAM 처리
        if upper.startswith('.PARAM'):
            param_text = stripped[6:].strip()  # .PARAM 이후의 텍스트 추출
            # param_text 또한 '이름=값' 쌍이므로 _parse_param_pairs 함수를 재사용해 분석합니다.
            pairs, _, _ = _parse_param_pairs(param_text)
            entries = [ParamEntry(name=k, value=v) for k, v in pairs.items()]
            
            # 현재 스코프(LIB 블록 내부인지, 문서 외부 전역인지)에 맞게 객체 삽입
            if in_lib_block and current_lib is not None:
                current_lib.params.extend(entries)
            else:
                lib_file.global_params.extend(entries)
            pending_comments = []
            i += 1
            continue

        # [조건 5] 기타 .(dot) 으로 시작하는 지시어(Directive) 처리
        # 예: .temp, .global, .options 등 보존해야 할 설정 영역
        if stripped.startswith('.'):
            parts = stripped.split(None, 1)
            keyword = parts[0]
            directive_entry = DirectiveEntry(keyword=keyword, raw_text=stripped)
            
            if in_lib_block and current_lib is not None:
                current_lib.directives.append(directive_entry)
            else:
                lib_file.global_directives.append(directive_entry)
            i += 1
            continue

        # 그 외 처리되지 않은 줄은 무시하고 다음 줄로 이동
        i += 1

    # 파일 최상단 주석 (아직 lib block이나 모델에 귀속되지 않은 경우)
    if not lib_file.lib_blocks and pending_comments:
        lib_file.leading_comments = pending_comments
    
    return lib_file
