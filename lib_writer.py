"""
lib_writer.py
LibFile 객체를 Smart Spice 문법의 .lib 텍스트로 직렬화합니다.

긴 파라미터 줄은 + continuation line으로 자동 줄바꿈 (기본 80자).
"""
from data_model import LibFile, LibBlock, ModelEntry, ParamEntry, DirectiveEntry

_LINE_WIDTH = 80


def _format_params(params: dict, indent: str = "+ ", open_paren: bool = False, close_paren: bool = False) -> list:
    """
    파라미터 딕셔너리를 Smart Spice 문법에 맞는 문자열 리스트(멀티 라인)로 변환합니다.
    
    특징:
    1. 긴 줄을 `_LINE_WIDTH`(기본 80자) 기준으로 여러 줄로 자동 개행합니다.
    2. 파싱 때 기록해둔 `open_paren`과 `close_paren` 값을 참조하여
       원형처럼 괄호로 파라미터를 감싸서 출력합니다.
       
    예:
      + (vth0=0.45 tox=1.2e-8
      + level=3)
    """
    if not params:
        # 파라미터가 없는데 괄호만 있을 수도 있음
        if open_paren and close_paren:
            return [indent + "()"]
        elif open_paren:
            return [indent + "("]
        elif close_paren:
            return [indent + ")"]
        return []

    items = [f"{k}={v}" for k, v in params.items()]
    lines = []
    
    current = indent
    if open_paren:
        current += "("
        
    for i, item in enumerate(items):
        is_last = (i == len(items) - 1)
        
        # 마지막 항목이고 close_paren이 True이면 ')' 붙임
        token = item
        if is_last and close_paren:
            token += ")"
            
        if len(current) + len(token) + 1 > _LINE_WIDTH and current.strip() not in ('+', '+(', '+ ('):
            lines.append(current.rstrip())
            current = indent + token + ' '
        else:
            current += token + ' '
            
    if current.strip() and current.strip() not in ('+', '+(', '+ ('):
        lines.append(current.rstrip())
    return lines


def _write_param_entries(entries: list) -> list:
    """
    ParamEntry 객체 리스트를 `.PARAM` 명령어 라인들로 직렬화합니다.
    여러 개의 파라미터가 있을 경우 80자를 넘지 않게 `+` continuation 라인으로 이어서 생성합니다.
    """
    if not entries:
        return []
    lines = []
    current = ".PARAM "
    for e in entries:
        token = f"{e.name}={e.value} "
        if len(current) + len(token) > _LINE_WIDTH and current.strip() != '.PARAM':
            lines.append(current.rstrip())
            current = "+ " + token
        else:
            current += token
    if current.strip() not in ('.PARAM', '+'):
        lines.append(current.rstrip())
    return lines


def write_lib(lib_file: LibFile) -> str:
    """
    메모리에 올려진 `LibFile` 구조 객체를 다시 순수 텍스트(Smart Spice 문법 문자열) 로 변환(직렬화)합니다.
    트리 순서(전역 영역 -> 각 LIB 블록 -> 내부 모델)를 그대로 따르며, 공백이나 주석의 위치도 본래대로 복구합니다.
    """
    out = []

    # 파일 선두 주석
    for c in lib_file.leading_comments:
        out.append(c)

    # 전역 .PARAM
    if lib_file.global_params:
        out.extend(_write_param_entries(lib_file.global_params))
        out.append('')

    # 전역 기타 directive
    if lib_file.global_directives:
        for d in lib_file.global_directives:
            out.append(d.raw_text)
        out.append('')

    # LIB 블록들
    for lb in lib_file.lib_blocks:
        # LIB 앞 주석
        for c in lb.leading_comments:
            out.append(c)

        out.append(f".LIB {lb.name}")

        # LIB 내 .PARAM
        if lb.params:
            out.extend(_write_param_entries(lb.params))

        # LIB별 기타 directive
        if lb.directives:
            for d in lb.directives:
                out.append(d.raw_text)
            out.append('')

        # MODEL 엔트리들
        for model in lb.models:
            for c in model.comment_lines:
                out.append(c)
            # 괄호 포함 모델 쓰기
            model_header = f".MODEL {model.name} {model.model_type}"
            out.append(model_header)
            
            if model.params or model.open_paren or model.close_paren:
                out.extend(_format_params(model.params, indent="+ ", open_paren=model.open_paren, close_paren=model.close_paren))

        out.append(f".ENDL {lb.name}")
        out.append('')

    return '\n'.join(out)


def save_lib(lib_file: LibFile, filepath: str = None) -> str:
    """
    메인 애플리케이션 등에서 호출하는 저장 API입니다.
    텍스트로 변환된 `LibFile`을 실제 로컬 파일 시스템의 경로(`filepath`)에 저장합니다.
    """
    path = filepath or lib_file.filepath
    if not path:
        raise ValueError("저장할 파일 경로가 지정되지 않았습니다.")
    content = write_lib(lib_file)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    return path
