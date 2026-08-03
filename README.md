# 出口单证生成平台（Streamlit 版）

把两个 Flask 小工具合并成一个网页，侧边栏切换不同工厂：

- **东莞致嘉金属**：上传发票 → 生成 申报要素(Word) / 成交确认书(Word) / 出口报关单(Excel)
- **宜章仁创**：上传装箱单/购销合同 → 生成 申报要素(Word) / 出口报关单(Excel)

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

浏览器打开 http://localhost:8501 即可使用。

## 发布到公网（Streamlit Community Cloud，免费）

### 1. 把项目推到 GitHub

1. 打开 https://github.com/new 新建一个仓库（例如 `export-docs-app`，选 Public）。
2. 在终端上传：

```bash
cd /Users/haoya/Desktop/export_docs_web
git init
git add .
git commit -m "init: 出口单证生成平台"
git branch -M main
git remote add origin https://github.com/<你的用户名>/export-docs-app.git
git push -u origin main
```

### 2. 连接到 Streamlit 云端

1. 用 GitHub 账号登录 https://share.streamlit.io
2. 点 **New app** → **Create app**
3. 选择仓库、分支（main），Main file 填 `app.py`，点击 **Deploy**
4. 等 1~2 分钟部署完成后会得到一个地址：
   `https://<你的用户名>-export-docs-app-main.streamlit.app`

把这个网址发给同事，任何人点开就能用，随时访问（☑ 支持手机浏览器）。

> 提示：免费版网页闲置一段时间后会“休眠”，有人打开时自动唤醒，稍等一下即可。
> 默认是公开的（任何有链接的人都能访问）。如果需要加密码，可以把
> `app.py` 顶部加上一段基于 `st.secrets` 的简单密码验证再部署。

## 更换模板/公司信息

- 两个工厂的 Word/Excel 模板放在 `static/zhijia/` 和 `static/renchuang/`，
  直接替换同名文件即可（保持文件名和占位符不变）。
- 增加新工厂：参考 `zhijia.py` / `renchuang.py` 的写法，在 `app.py` 的
  `FACTORIES` 和 `FUNCTION_MAP` 注册一项即可。