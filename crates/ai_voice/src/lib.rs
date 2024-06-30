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

        let voice_preset_names = ai_voice.voice_preset_names()?;
        println!("{:?}", voice_preset_names);
        println!("{:?}", ai_voice.master_control()?);

        println!("{:?}", ai_voice.voice_preset(&voice_preset_names[0])?);

        ai_voice.set_text("こんにちは")?;

        ai_voice.play()?;

        Ok(())
    }
}
