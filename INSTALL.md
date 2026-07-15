# CodexImage-PPT installation

## 给接收附件的 GPT

请将附件安装为本地 Codex Skill。先阅读压缩包中 `codeximage-ppt/INSTALL.md`，再把完整的 `codeximage-ppt` 文件夹安装到 Codex Skills 目录；不要只提取 `SKILL.md`。安装后请验证 Skill 结构，并告诉我是否需要重新启动或新建任务才能识别 `$codeximage-ppt`。

This archive contains one Codex skill. Its user-facing name is `CodexImage-PPT`; its valid internal identifier and folder name are `codeximage-ppt`.

## Instructions for the receiving GPT or Codex agent

1. Extract the archive and locate the `codeximage-ppt` folder containing `SKILL.md`.
2. Verify that `SKILL.md` declares `name: codeximage-ppt` and that `agents/openai.yaml` uses `$codeximage-ppt` in `default_prompt`.
3. Install the complete folder at `$CODEX_HOME/skills/codeximage-ppt`. If `CODEX_HOME` is unset, use the platform's normal Codex skills directory, commonly `~/.codex/skills/codeximage-ppt`.
4. Do not overwrite an existing skill with the same identifier without the user's approval.
5. Validate the installed folder with the environment's Skill validator when available, then start a new task or restart/reload Codex if the skill is not discovered immediately.

Uploading this ZIP as an ordinary chat attachment does not by itself prove that the chat product can install local Codex skills. The receiving environment must expose local Skill installation or use the OpenAI Skills API, which accepts a single ZIP skill upload.

## Runtime requirements

Installing the Skill and running every workflow are separate checks.

- Python dependencies are listed in `requirements.txt`.
- OCR requires the external Tesseract executable in addition to the `pytesseract` Python package. Chinese OCR also requires the `chi_sim` language data. If Tesseract is unavailable, the included decomposition script degrades to OCR-disabled output and readable text must be verified or reconstructed by another reliable method.
- Final rendered-deck QA requires Microsoft PowerPoint or an equivalent faithful renderer.
- Mandatory generated visual-core steps require an available ImageGen capability and its local transparency-removal workflow.

No personal credentials, API keys, or machine-specific absolute paths are included in this package.
