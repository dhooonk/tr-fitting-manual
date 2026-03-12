"""
data_model.py
Smart Spice LIB 파일의 데이터 모델 정의
"""
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class ParamEntry:
    """
    .PARAM 선언의 단일 변수를 나타내는 데이터 클래스입니다.
    예: .PARAM tox_val=1.2e-8
    여기서 tox_val은 name이 되고, 1.2e-8은 value가 됩니다.
    """
    name: str
    value: str  # 파라미터 값 (단순 숫자, 수식 {vth_offset * 1.1} 등 모두 문자열로 보관)


@dataclass
class DirectiveEntry:
    """
    정의된 구조(.PARAM, .MODEL, .LIB, .ENDL) 이외의 기타 .(dot) 시작 지시어 명령줄을 저장하는 클래스입니다.
    원본 파일에 존재하던 여타 설정값들을 손실 없이 그대로 저장하고 다시 내보내기 위해 사용됩니다.
    예:
      .temp 27
      .global vdd
    """
    keyword: str       # 분석된 지시어 첫번째 단어 (예: ".temp")
    raw_text: str      # 줄 전체 원문 텍스트 (예: ".temp 27")


@dataclass
class ModelEntry:
    """
    .MODEL 명령 하나와 그 하위에 종속된 파라미터 목록을 나타내는 데이터 클래스입니다.
    예:
      * 이것은 NMOS 모델입니다 (comment_lines)
      .MODEL NMOS_1 NMOS
      + (VTH0=0.45 TOX=1.2e-8)
      
    이 구조에서는 원래 모델에 괄호 ( )가 씌워져 있었는지를 파악하여 다시 저장할 때 원본 포맷을 유지합니다.
    """
    name: str                            # 모델명 (예: "NMOS_1")
    model_type: str                      # 모델 타입 (예: "NMOS" 또는 "PMOS")
    params: OrderedDict = field(default_factory=OrderedDict)
    # params: { '파라미터명': '값', ... } - 순서를 유지하는 딕셔너리
    comment_lines: List[str] = field(default_factory=list)  # 모델 선언 윗부분에 적혀있던 주석 목록
    open_paren: bool = False             # 타입명 뒤 또는 파라미터 리스트 맨 앞에 여는 형식의 괄호 '(' 가 있었는지 여부
    close_paren: bool = False            # 파라미터 리스트 맨 끝에 닫는 형식의 괄호 ')' 가 있었는지 여부

    def copy(self) -> "ModelEntry":
        return ModelEntry(
            name=self.name,
            model_type=self.model_type,
            params=OrderedDict(self.params),
            comment_lines=list(self.comment_lines),
            open_paren=self.open_paren,
            close_paren=self.close_paren
        )


@dataclass
class LibBlock:
    """
    .LIB 부터 .ENDL 까지 묶이는 하나의 독립된 라이브러리 블록을 나타냅니다.
    스마트 스파이스 모델 파일은 여러 개의 LIB 블록 내부로 파일 내용을 그룹화합니다.
    """
    name: str                                            # LIB 블록 이름
    models: List[ModelEntry] = field(default_factory=list) # 블록 내부에 속한 .MODEL 배열
    params: List[ParamEntry] = field(default_factory=list) # 블록 내부에서 선언된 .PARAM 변수 배열
    directives: List[DirectiveEntry] = field(default_factory=list) 
    # .PARAM 외에 블록 내부에 속한 여러 지시어 명령어들
    leading_comments: List[str] = field(default_factory=list) 
    # .LIB 선언부 윗부분에 적혀있던 블록 주석들

    def find_model(self, name: str) -> Optional[ModelEntry]:
        for m in self.models:
            if m.name.upper() == name.upper():
                return m
        return None


@dataclass
class LibFile:
    """
    파싱된 하나의 .lib 파일 전체를 루트 레벨에서 관리하는 최상단 데이터 구조입니다.
    전역 변수(블록 밖에 있는 값들)와 개별 LIB 블록 목록을 포함합니다.
    """
    filepath: str = ""                                               # 현재 파싱/저장된 원본 파일 경로
    global_params: List[ParamEntry] = field(default_factory=list)    # 전역 레벨에 설정된 .PARAM 들
    global_directives: List[DirectiveEntry] = field(default_factory=list)
    # 전역 레벨에 설정된 기타 지시어 명령들 (.temp 등)
    lib_blocks: List[LibBlock] = field(default_factory=list)         # 파일 안의 모든 .LIB 블록들 모음
    leading_comments: List[str] = field(default_factory=list)
    # 가장 윗단 파일 첫머리에 적혀있는 요약/설명 주석들 모음

    def find_lib(self, name: str) -> Optional[LibBlock]:
        for lb in self.lib_blocks:
            if lb.name.upper() == name.upper():
                return lb
        return None

    def all_params(self) -> List[ParamEntry]:
        """전역 + 모든 LIB 블록의 param 합계"""
        result = list(self.global_params)
        for lb in self.lib_blocks:
            result.extend(lb.params)
        return result
