use ai_voice::*;

mod ai_voice;

#[cfg(test)]
mod tests {
    use std::hint;

    use anyhow::Result;

    use super::*;

    #[test]
    fn main() -> Result<()> {
        let ai_voice = AiVoice::new()?;

        ai_voice.start_host()?;
        ai_voice.connect()?;

        let voice_names = ai_voice.voice_names()?;

        for voice_name in voice_names {
            println!("{}", voice_name);
            ai_voice.set_current_voice_preset_name(&voice_name)?;

            let voice_preset_names = ai_voice.voice_preset_names()?;
            println!("{:?}", voice_preset_names);
            println!("{:?}", ai_voice.voice_preset(&voice_preset_names[0])?);

            ai_voice.set_text("こんにちは")?;

            println!("play_time: {}", ai_voice.play_time()?);

            ai_voice.play()?;

            while ai_voice.status()? != HostStatus::Idle {
                hint::spin_loop();
            }
        }

        Ok(())
    }
}
