# ai_voice-rs

A.I.VOICE Editor API の非公式Rustバインディング

## 使用方法

`Cargo.toml`に以下を追加してください：

```toml
[dependencies]
ai_voice = { git = "https://github.com/nusu-github/ai_voice-rs" }
```

基本的な使用例：

```rust
use anyhow::Result;
use ai_voice::AiVoice;

fn main() -> Result<()> {
    let ai_voice = AiVoice::new()?;

    // A.I.VOICEを起動
    ai_voice.start_host()?;
    ai_voice.connect()?;

    // ボイスとパラメータを設定
    ai_voice.set_current_voice_preset_name("琴葉 茜")?;

    // 音声を生成
    ai_voice.set_text("こんにちは")?;
    ai_voice.play()?;

    // 再生が終わるまで待機
    while ai_voice.status()? != HostStatus::Idle {
        std::hint::spin_loop();
    }

    // A.I.VOICEから切断
    ai_voice.disconnect()?;
    ai_voice.terminate_host()?;

    Ok(())
}
```

## APIドキュメント

詳細なAPIドキュメントについては、プロジェクトディレクトリで`cargo doc --open`を実行してください。

## 依存クレート

このプロジェクトは以下の依存クレートを使用しています：

- `anyhow`: エラーハンドリング
- `serde`: 構造体のシリアライズとデシリアライズ
- `serde_json`: JSONのシリアライズとデシリアライズ
- `windows-rs`: Windows APIバインディング

## ライセンス

[Apache License, Version 2.0](LICENSE)

## 貢献

貢献を歓迎いたします！プルリクエストを気軽に送ってください。

## 参考リンク

- [A.I.VOICE 公式サイト](https://aivoice.jp/)
- [A.I.VOICE Editor ドキュメント A.I.VOICE Editor API](https://aivoice.jp/manual/editor/api.html)

## 免責事項

- 本ライブラリーは、株式会社エーアイ様、その他関係者とは一切関係がありません。
- 「A.I.VOICE」、「琴葉 茜」は株式会社エーアイの登録商標です。

## 注意事項

A.I.VOICEの使用にあたっては、A.I.VOICEの利用規約を遵守してください。本ライブラリを使用して生成されたコンテンツの利用に関しては、ユーザー自身の責任において行ってください。

A.I.VOICEの詳細な使用方法や最新の情報については、上記の公式サイトやAPIドキュメントを参照してください。
