"""
photo_dnd_component
--------------------
真正的 Streamlit 雙向自訂元件（純靜態 HTML/JS，不需要 Node/React 建置流程）。

跟 app.py 原本用「隱藏 st.text_input + 模擬事件」硬湊雙向溝通的做法不同，這裡改用
Streamlit 官方的元件通訊協定（window.postMessage + streamlit:setComponentValue），
好處是：Python 端資料改變時，前端的 iframe 不會被整個砍掉重建，畫面不會閃爍、
輸入文字時游標也不會被打斷。

用法：
    from photo_dnd_component import photo_dnd

    result = photo_dnd(
        photos={photo_key: base64_jpeg_str, ...},
        assignments={"01": ["Pxxxx", ...], "02": [], ...},
        descriptions={"01": "文字敘述...", ...},
        row_keys=["01", "02", ...],
        key="photo_dnd_widget",
    )
    if result:
        # result = {"assignments": {...}, "descriptions": {...}}
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


def photo_dnd(photos, assignments, descriptions, row_keys, key=None):
    """
    photos: dict，photo_key -> base64 編碼的 jpeg 縮圖字串（不含 data:image/... 前綴）
    assignments: dict，row_key -> [photo_key, ...]
    descriptions: dict，row_key -> 文字敘述字串
    row_keys: list[str]，例如 ["01", "02", "03"]
    key: Streamlit widget key

    回傳：
        None（使用者還沒跟這個元件互動過），或
        {"assignments": {...}, "descriptions": {...}}（使用者拖放或編輯文字之後的最新狀態）
    """
    return _component_func(
        photos=photos,
        assignments=assignments,
        descriptions=descriptions,
        rowKeys=row_keys,
        key=key,
        default=None,
    )
