# coding: utf-8
"""出口单证生成平台 - Streamlit 入口
侧边栏切换不同工厂/公司，进入对应的单证生成页面。
"""
import streamlit as st

import renchuang
import zhijia

st.set_page_config(
    page_title="出口单证生成平台",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="expanded",
)

FACTORIES = {
    "致凯金属 · 自动文档生成 (五金冲压模具/检具)": {"page": "zhijia", "desc": "上传发票 → 生成 申报要素 / 成交确认书 / 出口报关单"},
    "宜章仁创 · 液晶显示屏单证": {"page": "renchuang", "desc": "上传装箱单/购销合同 → 生成 申报要素 / 出口报关单"},
}

FUNCTION_MAP = {
    "zhijia": zhijia.render,
    "renchuang": renchuang.render,
}


def main():
    st.sidebar.title("📄 出口单证平台")
    st.sidebar.caption("选择工厂以切换到对应的生成工具")

    labels = list(FACTORIES.keys())
    choice = st.sidebar.radio("选择工厂", labels, label_visibility="collapsed")

    meta = FACTORIES[choice]
    st.sidebar.divider()
    st.sidebar.info(meta["desc"])

    with st.expander("当前页说明", expanded=False):
        st.write(f"当前正在使用：**{choice}**")
        st.write("页面会保留本次生成的下载按钮，重新生成会自动替换。")

    FUNCTION_MAP[meta["page"]]()


if __name__ == "__main__":
    main()