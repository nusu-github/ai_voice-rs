use std::{cmp::PartialEq, ffi::c_void, sync::Arc};

use anyhow::Result;
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

#[allow(non_snake_case)]
#[derive(Serialize, Deserialize, Debug)]
pub struct MasterControl {
    Volume: f32,
    Speed: f32,
    Pitch: f32,
    PitchRange: f32,
    MiddlePause: u16,
    LongPause: u16,
    SentencePause: u16,
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
        unsafe { self.control.IsInitialized() }
            .map_err(|err| err.into())
            .map(|x| x.as_bool())
    }

    pub fn start_host(&self) -> Result<()> {
        unsafe { self.control.StartHost() }.map_err(|err| err.into())
    }

    pub fn terminate_host(&self) -> Result<()> {
        unsafe { self.control.TerminateHost() }.map_err(|err| err.into())
    }

    pub fn connect(&self) -> Result<()> {
        unsafe { self.control.Connect() }.map_err(|err| err.into())
    }

    pub fn disconnect(&self) -> Result<()> {
        unsafe { self.control.Disconnect() }.map_err(|err| err.into())
    }

    pub fn version(&self) -> Result<String> {
        unsafe { self.control.Version() }
            .map_err(|err| err.into())
            .map(|x| x.to_string())
    }

    pub fn status(&self) -> Result<HostStatus> {
        let host_status: ai_voice_sys::HostStatus = unsafe { self.control.Status() }?;

        match host_status {
            ai_voice_sys::HostStatus(0) => Ok(HostStatus::NotRunning),
            ai_voice_sys::HostStatus(1) => Ok(HostStatus::NotConnected),
            ai_voice_sys::HostStatus(2) => Ok(HostStatus::Idle),
            ai_voice_sys::HostStatus(3) => Ok(HostStatus::Busy),
            _ => anyhow::bail!("Unknown host status"),
        }
    }

    pub fn master_control(&self) -> Result<MasterControl> {
        let master_control: BSTR = unsafe { self.control.MasterControl() }?;
        serde_json::from_str(master_control.to_string().as_str()).map_err(|err| err.into())
    }

    pub fn set_master_control(&self, master_control: &MasterControl) -> Result<()> {
        let master_control = MasterControl {
            Volume: master_control.Volume.clamp(0.0, 5.0),
            Speed: master_control.Speed.clamp(0.0, 4.0),
            Pitch: master_control.Pitch.clamp(0.0, 2.0),
            PitchRange: master_control.PitchRange.clamp(0.0, 2.0),
            MiddlePause: master_control.MiddlePause.clamp(0, 500),
            LongPause: master_control.LongPause.clamp(0, 2000),
            SentencePause: master_control.SentencePause.clamp(0, 10000),
        };

        let master_control = serde_json::to_string(&master_control)?;
        unsafe { self.control.SetMasterControl(&BSTR::from(master_control)) }
            .map_err(|err| err.into())
    }

    pub fn text(&self) -> Result<String> {
        unsafe { self.control.Text() }
            .map_err(|err| err.into())
            .map(|x| x.to_string())
    }

    pub fn set_text(&self, value: &str) -> Result<()> {
        unsafe { self.control.SetText(&BSTR::from(value)) }.map_err(|err| err.into())
    }

    pub fn text_selection_start(&self) -> Result<i32> {
        unsafe { self.control.TextSelectionStart() }.map_err(|err| err.into())
    }

    pub fn set_text_selection_start(&self, value: i32) -> Result<()> {
        unsafe { self.control.SetTextSelectionStart(value) }.map_err(|err| err.into())
    }

    pub fn text_selection_length(&self) -> Result<i32> {
        unsafe { self.control.TextSelectionLength() }.map_err(|err| err.into())
    }

    pub fn set_text_selection_length(&self, value: i32) -> Result<()> {
        unsafe { self.control.SetTextSelectionLength(value) }.map_err(|err| err.into())
    }

    pub fn text_edit_mode(&self) -> Result<TextEditMode> {
        let text_edit_mode: ai_voice_sys::TextEditMode = unsafe { self.control.TextEditMode() }?;

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

        unsafe { self.control.SetTextEditMode(text_edit_mode) }.map_err(|err| err.into())
    }

    pub fn play(&self) -> Result<()> {
        unsafe { self.control.Play() }.map_err(|err| err.into())
    }

    pub fn stop(&self) -> Result<()> {
        unsafe { self.control.Stop() }.map_err(|err| err.into())
    }

    pub fn save_audio_to_file(&self, path: &str) -> Result<()> {
        unsafe { self.control.SaveAudioToFile(&BSTR::from(path)) }.map_err(|err| err.into())
    }

    pub fn play_time(&self) -> Result<i64> {
        unsafe { self.control.GetPlayTime() }.map_err(|err| err.into())
    }

    pub fn list_count(&self) -> Result<i32> {
        unsafe { self.control.GetListCount() }.map_err(|err| err.into())
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
        unsafe { self.control.GetListSelectionCount() }.map_err(|err| err.into())
    }

    pub fn set_list_selection_index(&self, index: i32) -> Result<()> {
        unsafe { self.control.SetListSelectionIndex(index) }.map_err(|err| err.into())
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

        unsafe { self.control.SetListSelectionIndices(psa) }.map_err(|err| err.into())
    }

    pub fn set_list_selection_range(&self, startindex: i32, length: i32) -> Result<()> {
        unsafe { self.control.SetListSelectionRange(startindex, length) }.map_err(|err| err.into())
    }

    pub fn add_list_item(&self, voice_preset_name: &str, text: &str) -> Result<()> {
        unsafe {
            self.control
                .AddListItem(&BSTR::from(voice_preset_name), &BSTR::from(text))
        }
        .map_err(|err| err.into())
    }

    pub fn insert_list_item(&self, voice_preset_name: &str, text: &str) -> Result<()> {
        unsafe {
            self.control
                .InsertListItem(&BSTR::from(voice_preset_name), &BSTR::from(text))
        }
        .map_err(|err| err.into())
    }

    pub fn remove_list_item(&self) -> Result<()> {
        unsafe { self.control.RemoveListItem() }.map_err(|err| err.into())
    }

    pub fn clear_list_items(&self) -> Result<()> {
        unsafe { self.control.ClearListItems() }.map_err(|err| err.into())
    }

    pub fn list_voice_preset(&self) -> Result<String> {
        unsafe { self.control.GetListVoicePreset() }
            .map_err(|err| err.into())
            .map(|x| x.to_string())
    }

    pub fn set_list_voice_preset(&self, voice_preset_name: &str) -> Result<()> {
        unsafe {
            self.control
                .SetListVoicePreset(&BSTR::from(voice_preset_name))
        }
        .map_err(|err| err.into())
    }

    pub fn list_sentence(&self) -> Result<String> {
        unsafe { self.control.GetListSentence() }
            .map_err(|err| err.into())
            .map(|x| x.to_string())
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
        unsafe { self.control.CurrentVoicePresetName() }
            .map_err(|err| err.into())
            .map(|x| x.to_string())
    }

    pub fn set_current_voice_preset_name(&self, preset_name: &str) -> Result<()> {
        unsafe {
            self.control
                .SetCurrentVoicePresetName(&BSTR::from(preset_name))
        }
        .map_err(|err| err.into())
    }

    pub fn voice_preset(&self, preset_name: &str) -> Result<String> {
        unsafe { self.control.GetVoicePreset(&BSTR::from(preset_name)) }
            .map_err(|err| err.into())
            .map(|x| x.to_string())
    }

    pub fn set_voice_preset(&self, json: &str) -> Result<()> {
        unsafe { self.control.SetVoicePreset(&BSTR::from(json)) }.map_err(|err| err.into())
    }

    pub fn add_voice_preset(&self, json: &str) -> Result<()> {
        unsafe { self.control.AddVoicePreset(&BSTR::from(json)) }.map_err(|err| err.into())
    }

    pub fn reload_voice_presets(&self) -> Result<()> {
        unsafe { self.control.ReloadVoicePresets() }.map_err(|err| err.into())
    }

    pub fn reload_phrase_dictionary(&self) -> Result<()> {
        unsafe { self.control.ReloadPhraseDictionary() }.map_err(|err| err.into())
    }

    pub fn reload_word_dictionary(&self) -> Result<()> {
        unsafe { self.control.ReloadWordDictionary() }.map_err(|err| err.into())
    }

    pub fn reload_symbol_dictionary(&self) -> Result<()> {
        unsafe { self.control.ReloadSymbolDictionary() }.map_err(|err| err.into())
    }
}
