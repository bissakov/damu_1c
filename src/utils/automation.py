from time import time, sleep
from typing import (
    overload,
    Literal,
    Optional,
    Callable,
    Any,
    cast,
    Union,
    Generator,
    List,
    TypeVar,
    NewType,
    TypeAlias,
    Type,
    Dict,
)

from _ctypes import COMError
from pywinauto import WindowSpecification, Application
from pywinauto.controls.uia_controls import ButtonWrapper, EditWrapper, ListViewWrapper, ListItemWrapper
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.keyboard import send_keys

_WindowWindowSpecification = NewType("_WindowWindowSpecification", WindowSpecification)
_ButtonWindowSpecification = NewType("_ButtonWindowSpecification", WindowSpecification)
_CheckBoxWindowSpecification = NewType("_CheckBoxWindowSpecification", WindowSpecification)
_CustomWindowSpecification = NewType("_CustomWindowSpecification", WindowSpecification)
_DocumentWindowSpecification = NewType("_DocumentWindowSpecification", WindowSpecification)
_EditWindowSpecification = NewType("_EditWindowSpecification", WindowSpecification)
_ListWindowSpecification = NewType("_ListWindowSpecification", WindowSpecification)
_ListItemWindowSpecification = NewType("_ListItemWindowSpecification", WindowSpecification)
_PaneWindowSpecification = NewType("_PaneWindowSpecification", WindowSpecification)
_TabItemWindowSpecification = NewType("_TabItemWindowSpecification", WindowSpecification)
_TableWindowSpecification = NewType("_TableWindowSpecification", WindowSpecification)

_UIAWrapper = NewType("_UIAWrapper", UIAWrapper)
_ButtonWrapper = NewType("_ButtonWrapper", ButtonWrapper)
_CheckBoxWrapper = NewType("_CheckBoxWrapper", ButtonWrapper)
_UIACustomWrapper = NewType("_UIACustomWrapper", UIAWrapper)
_UIADocumentWrapper = NewType("_UIADocumentWrapper", UIAWrapper)
_EditWrapper = NewType("_EditWrapper", EditWrapper)
_ListViewWrapper = NewType("_ListViewWrapper", ListViewWrapper)
_ListItemWrapper = NewType("_ListItemWrapper", ListItemWrapper)
_UIAPaneWrapper = NewType("_UIAWPanerapper", UIAWrapper)
_UIATabItemWrapper = NewType("_UIATabItemWrapper", UIAWrapper)
_UIATableWrapper = NewType("_UIATableWrapper", ListViewWrapper)


UiaWindow = Union[_WindowWindowSpecification, _UIAWrapper]
UiaButton = Union[_ButtonWindowSpecification, _ButtonWrapper]
UiaCheckBox = Union[_CheckBoxWindowSpecification, _CheckBoxWrapper]
UiaCustom = Union[_CustomWindowSpecification, _UIACustomWrapper]
UiaDocument = Union[_DocumentWindowSpecification, _UIADocumentWrapper]
UiaEdit = Union[_EditWindowSpecification, _EditWrapper]
UiaList = Union[_ListWindowSpecification, _ListViewWrapper]
UiaListItem = Union[_ListItemWindowSpecification, _ListItemWrapper]
UiaPane = Union[_PaneWindowSpecification, _UIAPaneWrapper]
UiaTabItem = Union[_TabItemWindowSpecification, _UIATabItemWrapper]
UiaTable = Union[_TableWindowSpecification, _UIATableWrapper]


# UiaElement = Union[
#     WindowSpecification,
#     _UIAWrapper,
#     _ButtonWrapper,
#     _CheckBoxWrapper,
#     _UIACustomWrapper,
#     _UIADocumentWrapper,
#     _EditWrapper,
#     _ListViewWrapper,
#     _ListItemWrapper,
#     _UIAPaneWrapper,
#     _UIATabItemWrapper,
#     _UIATableWrapper,
# ]

UiaElement = TypeVar(
    "UiaElement",
    UiaWindow,
    UiaButton,
    UiaCheckBox,
    UiaCustom,
    UiaDocument,
    UiaEdit,
    UiaList,
    UiaListItem,
    UiaPane,
    UiaTabItem,
    UiaTable,
)


@overload
def child_win(parent: UiaElement, ctrl: Literal["Button"], title: Optional[str] = None, idx: int = 0) -> UiaButton: ...
@overload
def child_win(parent: UiaElement, ctrl: Literal["CheckBox"], title: Optional[str] = None, idx: int = 0) -> UiaCheckBox: ...
@overload
def child_win(parent: UiaElement, ctrl: Literal["Custom"], title: Optional[str] = None, idx: int = 0) -> UiaCustom: ...
@overload
def child_win(parent: UiaElement, ctrl: Literal["Document"], title: Optional[str] = None, idx: int = 0) -> UiaDocument: ...
@overload
def child_win(parent: UiaElement, ctrl: Literal["Edit"], title: Optional[str] = None, idx: int = 0) -> UiaEdit: ...
@overload
def child_win(parent: UiaElement, ctrl: Literal["List"], title: Optional[str] = None, idx: int = 0) -> UiaList: ...
@overload
def child_win(parent: UiaElement, ctrl: Literal["ListItem"], title: Optional[str] = None, idx: int = 0) -> UiaListItem: ...
@overload
def child_win(parent: UiaElement, ctrl: Literal["Pane"], title: Optional[str] = None, idx: int = 0) -> UiaPane: ...
@overload
def child_win(parent: UiaElement, ctrl: Literal["TabItem"], title: Optional[str] = None, idx: int = 0) -> UiaTabItem: ...
@overload
def child_win(parent: UiaElement, ctrl: Literal["Table"], title: Optional[str] = None, idx: int = 0) -> UiaTable: ...


def child_win(
    parent: UiaElement,
    ctrl: Literal["Button", "CheckBox", "Custom", "Document", "Edit", "List", "ListItem", "Pane", "TabItem", "Table"],
    title: Optional[str] = None,
    idx: int = 0,
):
    return parent.child_window(title=title, control_type=ctrl, found_index=idx)


def window(app: Application, title: str, regex: bool = False) -> UiaWindow:
    if regex:
        return app.window(title_re=title)
    else:
        return app.window(title=title)


def a(main_win: UiaWindow, action: Callable[[], None]) -> None:
    if not main_win.is_active():
        main_win.set_focus()
        main_win.wait(wait_for="active visible")

    action()


@overload
def iter_children(parent: UiaList) -> Generator[UiaListItem, None, None]: ...
@overload
def iter_children(parent: UiaTable) -> Generator[UiaCustom, None, None]: ...


def iter_children(parent: UiaElement):
    return parent.iter_children()


def children(parent: UiaList) -> List[UiaListItem]:
    return parent.children()


def wait_for(condition: Callable[[], bool], timeout: float, interval: float = 0.1) -> bool:
    start = time()
    while not condition():
        if time() - start > timeout:
            return False
        sleep(interval)
    return True


def click_type_keys(
    element: Union[UiaEdit, UiaCustom],
    keystrokes: Any,
    delay: float = 0.1,
    pause: float = 0.05,
    double: bool = False,
    cls: bool = True,
    ent: bool = False,
    spaces: bool = False,
    escape_chars: bool = False,
) -> None:
    if not isinstance(keystrokes, str):
        keystrokes = cast(str, str(keystrokes))

    if escape_chars:
        keystrokes = keystrokes.replace("\n", "{ENTER}").replace("(", "{(}").replace(")", "{)}")

    if cls:
        keystrokes = "{DELETE}" + keystrokes

    if ent:
        keystrokes = keystrokes + "{ENTER}+{TAB}"

    if double:
        element.double_click_input()
    else:
        element.click_input()
    sleep(delay)
    send_keys(keystrokes, pause=pause, with_spaces=spaces)


def check(checkbox: UiaCheckBox) -> None:
    if checkbox.get_toggle_state() == 0:
        checkbox.toggle()


def contains_text(element: UiaElement) -> bool:
    return any((inner.strip() for outer in element.texts() for inner in outer))


def text(element: UiaElement) -> str:
    return cast(str, element.window_text())


def text_to_float(txt: str, default: Optional[float] = None) -> float:
    try:
        res = float(txt.replace(",", "."))
        return res
    except ValueError as err:
        if isinstance(default, float):
            return default
        raise err


def get_full_text(element: UiaElement) -> str:
    txt = element.window_text().strip() if element.window_text() else ""

    for child in element.children():
        child_text = get_full_text(child)
        if child_text:
            txt += " " + child_text

    return txt.strip()


def print_element_tree(
    element: UiaElement, max_depth: Optional[int] = None, counters: Optional[Dict[str, int]] = None, depth: int = 0
) -> None:
    """
    :param element: UiaElement - Root element of the tree
    :param max_depth: Optional[int} = None - Max depth of the tree to print
    :param counters: Optional[Dict[str, int]] = None - Index count (**IMPORTANT! don't use directly**). Not accurate when max_depth is set
    :param depth: Optional[int] = 0 - Current depth of the tree (**IMPORTANT! don't use directly**)
    :return: None
    """

    if max_depth is not None:
        if not isinstance(max_depth, int) or max_depth < 0:
            raise ValueError("max_depth must be a non-negative integer or None")

    if counters is None:
        counters = {}

    element_ctrl = element.friendly_class_name()
    counters[element_ctrl] = counters.get(element_ctrl, 0) + 1
    element_idx = counters[element_ctrl] - 1

    element_repr = "▏   " * (depth + 1) + f"{element_ctrl}{element_idx} - {text(element)!r} - "

    try:
        element_repr += f"{element.rectangle()}"
        print(element_repr)
    except COMError:
        element_repr += "(COMError)"
        print(element_repr)
        return

    if max_depth is None or depth < max_depth:
        for child in iter_children(element):
            print_element_tree(child, max_depth, counters, depth + 1)
