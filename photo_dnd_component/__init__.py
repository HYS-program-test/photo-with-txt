"""
photo_dnd_component
--------------------
純粹「拖曳照片指派」用的自訂元件（純靜態 HTML/JS，不需要 Node/React 建置流程）。

這一版只負責拖曳照片這件事，文字敘述不放在這裡面 —— 因為文字輸入是連續性的動作，
只要牽扯到跟 Python 來回同步，就容易在正式部署環境（有網路延遲）上感覺卡頓、
打字被打斷。拖曳是一次性的動作（放開手才同步一次），風險小很多，所以只把拖曳
留在這個自訂元件裡；文字敘述請在 app.py 裡用原生的 st.text_area（最不會出問題）。

用法：
    from photo_dnd_component import photo_dnd

    result = photo_dnd(
        photos={photo_key: base64_jpeg_str, ...},
        assignments={"01": ["Pxxxx", ...], "02": [], ...},
        row_keys=["01", "02", ...],
        key="photo_dnd_widget",
    )
    if result:
        # result = {"assignments": {...}}
        ...
"""

import os

import streamlit.components.v1 as components

_COMPONENT_NAME = "photo_dnd"
_FRONTEND_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "frontend")

_component_func = components.declare_component(
    _COMPONENT_NAME,
    path=_FRONTEND_DIR,
)


def photo_dnd(photos, assignments, row_keys, key=None):
    """
    photos: dict，photo_key -> base64 編碼的 jpeg 縮圖字串（不含 data:image/... 前綴）
    assignments: dict，row_key -> [photo_key, ...]
    row_keys: list[str]，例如 ["01", "02", "03"]
    key: Streamlit widget key

    回傳：
        None（使用者還沒跟這個元件互動過），或
        {"assignments": {...}}（使用者拖放之後的最新狀態）
    """
    return _component_func(
        photos=photos,
        assignments=assignments,
        rowKeys=row_keys,
        key=key,
        default=None,
    )
