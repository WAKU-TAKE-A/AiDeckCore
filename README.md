# AiDeckCore / deck2pptx

AI に PowerPoint 資料を作らせるための、意味構造ベースの PPTX 生成ツールです。

このプロジェクトでは `Deck` model を正規表現として扱います。YAML や Markdown は入力形式のひとつにすぎません。AI には PowerPoint の座標を直接指定させず、スライドの意味構造を書かせてから `deck2pptx` で PPTX に変換します。

## まず読むもの

- 人間向け: この `README.md`
- AI 向け: `README_AI.md`
- 詳細なリリース確認: `docs/release-verification.md`
- ソース管理方針: `docs/source-baseline.md`

AI に資料作成を依頼するときは、まず `README_AI.md` を見せてください。

## 環境作成

PowerShell でリポジトリ直下に移動して、repo-local の仮想環境を作成します。

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -e . pytest python-pptx pyyaml pillow
```

通常の作業では、この `.venv` を使ってください。

```powershell
.\.venv\Scripts\Activate.ps1
```

## 動作確認

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

リリース前相当の確認を行う場合は、次を実行します。

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\verify_release.ps1
```

この確認には、テスト、AI 向けコマンド、YAML/Markdown からの PPTX 生成、失敗すべき入力の検証、fresh venv でのインストール確認、ソース hygiene 確認が含まれます。

LibreOffice がある環境では、PPTX から PDF への visual export も確認します。Windows で Chocolatey が使える場合は、必要に応じて以下でインストールできます。

```powershell
choco install libreoffice-fresh -y
```

## 基本的な使い方

Markdown から PPTX を作る例です。

```powershell
.\.venv\Scripts\python.exe -m deck2pptx validate examples\sample.md
.\.venv\Scripts\python.exe -m deck2pptx build examples\sample.md outputs\sample.pptx
```

YAML からも同じ Deck model に変換して PPTX を作れます。

```powershell
.\.venv\Scripts\python.exe -m deck2pptx validate examples\sample.deck.yaml
.\.venv\Scripts\python.exe -m deck2pptx build examples\sample.deck.yaml outputs\sample-yaml.pptx
```

Markdown や YAML の先頭（Front Matter）に以下の設定を追加することで、目次やタイポグラフィの制御が可能です。
- `toc: true` (目次スライドの自動生成)
- `indent: 2` (Markdownにおけるリストの字下げスペース数)
- `font_size_l0: 24` (レベルごとのフォントサイズの直接指定)

既存の PPTX テンプレートを利用して生成することも可能です。
まずテンプレートのレイアウト名とプレースホルダー名を確認します。

```powershell
.\.venv\Scripts\python.exe -m deck2pptx inspect-template template.pptx
```

確認した正確な名前を Markdown や YAML に指定し、テンプレートを指定して build します。
レイアウト名とプレースホルダー名は、大文字小文字を無視した先頭一致で探索します。PowerPoint が実スライド上で `Subtitle 2` のように番号を付ける場合は、`Subtitle` や `sub` のような先頭名で指定できます。プレースホルダーは、レイアウト上の名前から同じ idx の実スライドプレースホルダーへ対応付けます。

```powershell
.\.venv\Scripts\python.exe -m deck2pptx build examples\sample_advanced.md outputs\advanced.pptx --template template.pptx
```

入力ファイル内の画像パスは、入力ファイルのあるフォルダからの相対パスとして解決されます。

## AI に作業させるとき

AI には `README_AI.md` を見せて、次の流れを守らせてください。

1. `explain-spec --format json` で現在の仕様を確認する。
2. YAML または Markdown で資料の入力ファイルを作る。
3. `inspect --format json` で Deck model への変換結果を確認する。
4. `validate --format json` で構造エラーを確認する。
5. エラーがあれば structured error に従って修正する。
6. `build` で PPTX を生成する。

AI に PowerPoint の座標、テキストボックスの位置、図形の絶対サイズを直接決めさせないでください。このプロジェクトでは Deck model の意味構造を入力し、renderer が PPTX に変換します。

## 現在の対応範囲

対応済みの Deck 要素:

- Text
- BulletList
- Image
- Table
- Gallery
- Flow
- Comparison
- Timeline
- CodeBlock
- Tree

対応済みの入力形式:

- YAML
- Markdownでは以下をサポートしています。

- テキスト、箇条書き、テーブル
- 画像（`![alt](path)`）と連続画像の自動ギャラリーレイアウト
- フローチャート（`` `flow` `` ブロック）
- 比較マトリックス（`` `comparison` `` ブロック）
- タイムライン（`` `timeline` `` ブロック）
- コードブロック（`` `code python` `` ブロック）
- ツリー構造（`` `tree` `` ブロック）
- HTMLコメントを使ったレイアウトヒント (`<!-- layout="Name" -->`)
- 要素のプレースホルダー指定 (`<!-- place="Name" -->`)

未対応または次フェーズの候補:

- AsciiDoc adapter
- Natural Language adapter
- 本格的な template/theme system

## ライセンス

このプロジェクトは MIT License です。詳細は `LICENSE` を参照してください。

## 設計方針

- `Deck` model が正規の表現です。
- YAML は入力 adapter です。
- Markdown も入力 adapter です。
- 将来の AsciiDoc や Natural Language も、同じ Deck model を target にします。
- Renderer は YAML/Markdown/AsciiDoc/Natural Language を直接読みません。
- PowerPoint は renderer のひとつです。
