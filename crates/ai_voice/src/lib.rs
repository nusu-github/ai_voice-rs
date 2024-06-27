mod ai_voice;

#[cfg(test)]
mod tests {
    use anyhow::Result;

    use super::*;

    #[test]
    fn main() -> Result<()> {
        let ai_voice = ai_voice::AiVoice::new()?;

        ai_voice.start_host()?;

        ai_voice.connect()?;

        println!("{:?}", ai_voice.voice_names()?);
        println!("{:?}", ai_voice.voice_preset_names()?);
        println!("{:?}", ai_voice.master_control()?);

        ai_voice.set_text("こんにちは")?;

        ai_voice.play()?;

        Ok(())
    }
}
