use std::{cmp::PartialEq, ffi::c_void, sync::Arc};

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};
use windows::{
    core::BSTR,
    Win32::System::{Com::*, Ole::*, Variant::*},
};

use ai_voice_sys::{ITtsControl, TtsControl};

#[derive(Debug, PartialEq)]
pub enum HostStatus {
    Busy,
    Idle,
    NotConnected,
    NotRunning,
}

#[derive(Debug, PartialEq)]
pub enum TextEditMode {
    List,
    Text,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct MasterControl {
    pub volume: f32,
    pub speed: f32,
    pub pitch: f32,
    pub pitch_range: f32,
    pub middle_pause: u16,
    pub long_pause: u16,
    pub sentence_pause: u16,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct VoicePreset {
    pub preset_name: String,
    pub voice_name: String,
    pub volume: f64,
    pub speed: f64,
    pub pitch: f64,
    pub pitch_range: f64,
    pub middle_pause: i64,
    pub long_pause: i64,
    pub styles: Vec<Style>,
    pub merged_voice_container: MergedVoiceContainer,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct MergedVoiceContainer {
    pub base_pitch_voice_name: String,
    pub merged_voices: Vec<MergedVoice>,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct MergedVoice {
    pub voice_name: String,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct Style {
    pub name: String,
    pub value: f64,
}

#[derive(Clone)]
pub struct AiVoice {
    control: Arc<ITtsControl>,
}

impl Drop for AiVoice {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

impl AiVoice {
    pub fn new() -> Result<Self> {
        unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED).ok()?;

            let control: ITtsControl = CoCreateInstance(&TtsControl, None, CLSCTX_INPROC_SERVER)?;

            let mut host_name = BSTR::default();
            SafeArrayGetElement(
                control.GetAvailableHostNames()?,
                &0,
                &mut host_name as *mut BSTR as *mut c_void,
            )?;

            control.Initialize(&host_name)?;

            Ok(AiVoice {
                control: Arc::new(control),
            })
        }
    }

    pub fn is_initialized(&self) -> Result<bool> {
        Ok(unsafe { self.control.IsInitialized() }?.as_bool())
    }

    pub fn start_host(&self) -> Result<()> {
        Ok(unsafe { self.control.StartHost() }?)
    }

    pub fn terminate_host(&self) -> Result<()> {
        Ok(unsafe { self.control.TerminateHost() }?)
    }

    pub fn connect(&self) -> Result<()> {
        Ok(unsafe { self.control.Connect() }?)
    }

    pub fn disconnect(&self) -> Result<()> {
        Ok(unsafe { self.control.Disconnect() }?)
    }

    pub fn version(&self) -> Result<String> {
        Ok(unsafe { self.control.Version() }?.to_string())
    }

    pub fn status(&self) -> Result<HostStatus> {
        let host_status = unsafe { self.control.Status() }?;

        match host_status {
            ai_voice_sys::HostStatus(0) => Ok(HostStatus::NotRunning),
            ai_voice_sys::HostStatus(1) => Ok(HostStatus::NotConnected),
            ai_voice_sys::HostStatus(2) => Ok(HostStatus::Idle),
            ai_voice_sys::HostStatus(3) => Ok(HostStatus::Busy),
            _ => anyhow::bail!("Unknown host status"),
        }
    }

    pub fn master_control(&self) -> Result<MasterControl> {
        let master_control = unsafe { self.control.MasterControl() }?.to_string();
        serde_json::from_str(&master_control).with_context(|| "Failed to parse master control")
    }

    pub fn set_master_control(&self, master_control: &MasterControl) -> Result<()> {
        let master_control = MasterControl {
            volume: master_control.volume.clamp(0.0, 5.0),
            speed: master_control.speed.clamp(0.0, 4.0),
            pitch: master_control.pitch.clamp(0.0, 2.0),
            pitch_range: master_control.pitch_range.clamp(0.0, 2.0),
            middle_pause: master_control.middle_pause.clamp(0, 500),
            long_pause: master_control.long_pause.clamp(0, 2000),
            sentence_pause: master_control.sentence_pause.clamp(0, 10000),
        };

        let master_control = serde_json::to_string(&master_control)?;
        Ok(unsafe { self.control.SetMasterControl(&BSTR::from(master_control)) }?)
    }

    pub fn text(&self) -> Result<String> {
        Ok(unsafe { self.control.Text() }?.to_string())
    }

    pub fn set_text(&self, value: &str) -> Result<()> {
        Ok(unsafe { self.control.SetText(&BSTR::from(value)) }?)
    }

    pub fn text_selection_start(&self) -> Result<i32> {
        Ok(unsafe { self.control.TextSelectionStart() }?)
    }

    pub fn set_text_selection_start(&self, value: i32) -> Result<()> {
        Ok(unsafe { self.control.SetTextSelectionStart(value) }?)
    }

    pub fn text_selection_length(&self) -> Result<i32> {
        Ok(unsafe { self.control.TextSelectionLength() }?)
    }

    pub fn set_text_selection_length(&self, value: i32) -> Result<()> {
        Ok(unsafe { self.control.SetTextSelectionLength(value) }?)
    }

    pub fn text_edit_mode(&self) -> Result<TextEditMode> {
        let text_edit_mode = unsafe { self.control.TextEditMode() }?;

        match text_edit_mode {
            ai_voice_sys::TextEditMode(0) => Ok(TextEditMode::Text),
            ai_voice_sys::TextEditMode(1) => Ok(TextEditMode::List),
            _ => anyhow::bail!("Unknown text edit mode"),
        }
    }

    pub fn set_text_edit_mode(&self, mode: TextEditMode) -> Result<()> {
        let text_edit_mode = match mode {
            TextEditMode::Text => ai_voice_sys::TextEditMode(0),
            TextEditMode::List => ai_voice_sys::TextEditMode(1),
        };

        Ok(unsafe { self.control.SetTextEditMode(text_edit_mode) }?)
    }

    pub fn play(&self) -> Result<()> {
        Ok(unsafe { self.control.Play() }?)
    }

    pub fn stop(&self) -> Result<()> {
        Ok(unsafe { self.control.Stop() }?)
    }

    pub fn save_audio_to_file(&self, path: &str) -> Result<()> {
        Ok(unsafe { self.control.SaveAudioToFile(&BSTR::from(path)) }?)
    }

    pub fn play_time(&self) -> Result<i64> {
        Ok(unsafe { self.control.GetPlayTime() }?)
    }

    pub fn list_count(&self) -> Result<i32> {
        Ok(unsafe { self.control.GetListCount() }?)
    }

    pub fn list_selection_indices(&self) -> Result<Vec<i32>> {
        let indices = unsafe { self.control.GetListSelectionIndices() }?;

        let lob = unsafe { SafeArrayGetLBound(indices, 1) }?;
        let upb = unsafe { SafeArrayGetUBound(indices, 1) }?;

        let mut selection_indices = Vec::with_capacity((lob.abs() + upb.abs() + 1) as usize);
        for i in lob..upb + 1 {
            let data = unsafe {
                let mut result__ = i32::default();
                SafeArrayGetElement(indices, &i, &mut result__ as *mut i32 as *mut _)?;
                result__
            };

            selection_indices.push(data);
        }

        Ok(selection_indices)
    }

    pub fn list_selection_count(&self) -> Result<i32> {
        Ok(unsafe { self.control.GetListSelectionCount() }?)
    }

    pub fn set_list_selection_index(&self, index: i32) -> Result<()> {
        Ok(unsafe { self.control.SetListSelectionIndex(index) }?)
    }

    pub fn set_list_selection_indices(&self, indices: Vec<String>) -> Result<()> {
        let rgsabound = vec![SAFEARRAYBOUND {
            cElements: indices.len() as u32,
            lLbound: 0,
        }];
        let psa = unsafe { SafeArrayCreate(VT_BSTR, 0, rgsabound.as_ptr()) };

        for (i, elem) in indices.iter().enumerate() {
            unsafe {
                SafeArrayPutElement(
                    psa,
                    &(i as i32),
                    std::mem::transmute_copy(&BSTR::from(elem)),
                )
            }?;
        }

        Ok(unsafe { self.control.SetListSelectionIndices(psa) }?)
    }

    pub fn set_list_selection_range(&self, startindex: i32, length: i32) -> Result<()> {
        Ok(unsafe { self.control.SetListSelectionRange(startindex, length) }?)
    }

    pub fn add_list_item(&self, voice_preset_name: &str, text: &str) -> Result<()> {
        Ok(unsafe {
            self.control
                .AddListItem(&BSTR::from(voice_preset_name), &BSTR::from(text))
        }?)
    }

    pub fn insert_list_item(&self, voice_preset_name: &str, text: &str) -> Result<()> {
        Ok(unsafe {
            self.control
                .InsertListItem(&BSTR::from(voice_preset_name), &BSTR::from(text))
        }?)
    }

    pub fn remove_list_item(&self) -> Result<()> {
        Ok(unsafe { self.control.RemoveListItem() }?)
    }

    pub fn clear_list_items(&self) -> Result<()> {
        Ok(unsafe { self.control.ClearListItems() }?)
    }

    pub fn list_voice_preset(&self) -> Result<String> {
        Ok(unsafe { self.control.GetListVoicePreset() }?.to_string())
    }

    pub fn set_list_voice_preset(&self, voice_preset_name: &str) -> Result<()> {
        Ok(unsafe {
            self.control
                .SetListVoicePreset(&BSTR::from(voice_preset_name))
        }?)
    }

    pub fn list_sentence(&self) -> Result<String> {
        Ok(unsafe { self.control.GetListSentence() }?.to_string())
    }

    pub fn voice_names(&self) -> Result<Vec<String>> {
        let voice_names = unsafe { self.control.VoiceNames() }?;

        let lob = unsafe { SafeArrayGetLBound(voice_names, 1) }?;
        let upb = unsafe { SafeArrayGetUBound(voice_names, 1) }?;

        let mut voice_names_vec = Vec::with_capacity((lob.abs() + upb.abs() + 1) as usize);
        for i in lob..upb + 1 {
            let data = unsafe {
                let mut result__ = BSTR::default();
                SafeArrayGetElement(voice_names, &i, &mut result__ as *mut BSTR as *mut _)?;
                result__
            };

            voice_names_vec.push(data.to_string());
        }

        Ok(voice_names_vec)
    }

    pub fn voice_preset_names(&self) -> Result<Vec<String>> {
        let preset_names = unsafe { self.control.VoicePresetNames() }?;

        let lob = unsafe { SafeArrayGetLBound(preset_names, 1) }?;
        let upb = unsafe { SafeArrayGetUBound(preset_names, 1) }?;

        let mut preset_names_vec = Vec::with_capacity((lob.abs() + upb.abs() + 1) as usize);
        for i in lob..upb {
            let data = unsafe {
                let mut result__ = BSTR::default();
                SafeArrayGetElement(preset_names, &i, &mut result__ as *mut BSTR as *mut _)?;
                result__
            };

            preset_names_vec.push(data.to_string());
        }

        Ok(preset_names_vec)
    }

    pub fn current_voice_preset_name(&self) -> Result<String> {
        Ok(unsafe { self.control.CurrentVoicePresetName() }?.to_string())
    }

    pub fn set_current_voice_preset_name(&self, preset_name: &str) -> Result<()> {
        Ok(unsafe {
            self.control
                .SetCurrentVoicePresetName(&BSTR::from(preset_name))
        }?)
    }

    pub fn voice_preset(&self, preset_name: &str) -> Result<VoicePreset> {
        let voice_preset =
            unsafe { self.control.GetVoicePreset(&BSTR::from(preset_name)) }?.to_string();
        serde_json::from_str(&voice_preset).with_context(|| "Failed to parse voice preset")
    }

    pub fn set_voice_preset(&self, voice_preset: &VoicePreset) -> Result<()> {
        let json = serde_json::to_string(voice_preset)?;
        Ok(unsafe { self.control.SetVoicePreset(&BSTR::from(json)) }?)
    }

    pub fn add_voice_preset(&self, voice_preset: &VoicePreset) -> Result<()> {
        let json = serde_json::to_string(voice_preset)?;
        Ok(unsafe { self.control.AddVoicePreset(&BSTR::from(json)) }?)
    }

    pub fn reload_voice_presets(&self) -> Result<()> {
        Ok(unsafe { self.control.ReloadVoicePresets() }?)
    }

    pub fn reload_phrase_dictionary(&self) -> Result<()> {
        Ok(unsafe { self.control.ReloadPhraseDictionary() }?)
    }

    pub fn reload_word_dictionary(&self) -> Result<()> {
        Ok(unsafe { self.control.ReloadWordDictionary() }?)
    }

    pub fn reload_symbol_dictionary(&self) -> Result<()> {
        Ok(unsafe { self.control.ReloadSymbolDictionary() }?)
    }
}
