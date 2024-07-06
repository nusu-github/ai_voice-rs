use std::{cmp::PartialEq, ffi::c_void, sync::Arc};

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};
use windows::{
    core::BSTR,
    Win32::System::{Com::*, Ole::*, Variant::*},
};

use ai_voice_sys::{ITtsControl, TtsControl};

#[derive(Debug, PartialEq)]
#[doc = "ホストプログラムの状態"]
pub enum HostStatus {
    #[doc = "起動していない"]
    NotRunning,
    #[doc = "接続していない"]
    NotConnected,
    #[doc = "アイドル状態"]
    Idle,
    #[doc = "処理中"]
    Busy,
}

#[derive(Debug, PartialEq)]
#[doc = "テキスト入力形式"]
pub enum TextEditMode {
    #[doc = "テキスト形式"]
    Text,
    #[doc = "リスト形式"]
    List,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
#[doc = "マスターコントロール"]
pub struct MasterControl {
    #[doc = "ボリューム"]
    pub volume: f32,
    #[doc = "話速"]
    pub speed: f32,
    #[doc = "高さ"]
    pub pitch: f32,
    #[doc = "抑揚"]
    pub pitch_range: f32,
    #[doc = "短ポーズ(ms)"]
    pub middle_pause: u16,
    #[doc = "長ポーズ(ms)"]
    pub long_pause: u16,
    #[doc = "文末ポーズ(ms)"]
    pub sentence_pause: u16,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
#[doc = "ボイスプリセット"]
pub struct VoicePreset {
    #[doc = "ボイスプリセット名"]
    pub preset_name: String,
    #[doc = "ボイス名"]
    pub voice_name: String,
    #[doc = "ボリューム"]
    pub volume: f32,
    #[doc = "話速"]
    pub speed: f32,
    #[doc = "高さ"]
    pub pitch: f32,
    #[doc = "抑揚"]
    pub pitch_range: f32,
    #[doc = "短ポーズ(ms)"]
    pub middle_pause: u16,
    #[doc = "長ポーズ(ms)"]
    pub long_pause: u16,
    #[doc = "スタイル情報のリスト"]
    pub styles: Vec<Style>,
    #[doc = "フュージョン情報"]
    pub merged_voice_container: MergedVoiceContainer,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
#[doc = "フュージョン情報"]
pub struct MergedVoiceContainer {
    #[doc = "高さのベースとなるボイス名"]
    pub base_pitch_voice_name: String,
    #[doc = "フュージョンされたボイスのリスト"]
    pub merged_voices: Vec<MergedVoice>,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct MergedVoice {
    #[doc = "ボイス名"]
    pub voice_name: String,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct Style {
    /// スタイル名
    ///
    /// # 指定可能な値
    /// - J: 喜び（ハイテンション）
    /// - A: 怒り
    /// - S: 悲しみ（ローテンション）
    ///
    pub name: String,
    #[doc = "値"]
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

    /// APIが初期化されているかどうかを取得します。
    ///
    pub fn is_initialized(&self) -> Result<bool> {
        Ok(unsafe { self.control.IsInitialized() }?.as_bool())
    }

    /// ホストプログラムを起動します。
    ///
    pub fn start_host(&self) -> Result<()> {
        Ok(unsafe { self.control.StartHost() }?)
    }

    /// ホストプログラムを終了します。
    ///
    /// # 注意
    /// ホストプログラムでプロジェクト等が変更状態の場合、
    /// ホストプログラム上で確認メッセージが表示されます。
    ///
    pub fn terminate_host(&self) -> Result<()> {
        Ok(unsafe { self.control.TerminateHost() }?)
    }

    /// ホストプログラムと接続します。
    ///
    /// # 注意
    /// ホストプログラムへ接続後、10分間 API を介した操作が行われない状態が続くと
    /// 自動的に接続が解除されます。
    ///
    pub fn connect(&self) -> Result<()> {
        Ok(unsafe { self.control.Connect() }?)
    }

    /// ホストプログラムとの接続を解除します。
    ///
    pub fn disconnect(&self) -> Result<()> {
        Ok(unsafe { self.control.Disconnect() }?)
    }

    /// ホストプログラムのバージョンを取得します。
    ///
    /// # 戻り値
    /// ホストプログラムのバージョン文字列
    ///
    pub fn version(&self) -> Result<String> {
        Ok(unsafe { self.control.Version() }?.to_string())
    }

    /// ホストプログラムの状態を取得します。
    ///
    /// # 戻り値
    /// `HostStatus` 列挙型で表されるホストプログラムの状態
    ///
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

    /// マスターコントロールの現在の設定を取得します。
    ///
    /// # 戻り値
    /// `MasterControl` 構造体で表されるマスターコントロールの設定
    ///
    pub fn master_control(&self) -> Result<MasterControl> {
        let master_control = unsafe { self.control.MasterControl() }?.to_string();
        serde_json::from_str(&master_control).with_context(|| "Failed to parse master control")
    }

    /// マスターコントロールの設定を適用します。
    ///
    /// この関数は、入力値を適切な範囲内に制限します。
    ///
    /// # 引数
    /// * `master_control` - 適用する `MasterControl` 構造体
    ///
    pub fn apply_master_control(&self, master_control: &MasterControl) -> Result<()> {
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

    /// テキスト形式の入力テキストを取得します。
    ///
    /// # 戻り値
    /// 現在設定されているテキスト
    ///
    pub fn text(&self) -> Result<String> {
        Ok(unsafe { self.control.Text() }?.to_string())
    }

    /// テキスト形式の入力テキストを設定します。
    ///
    /// # 引数
    /// * `value` - 設定するテキスト
    ///
    pub fn set_text(&self, value: &str) -> Result<()> {
        Ok(unsafe { self.control.SetText(&BSTR::from(value)) }?)
    }

    /// テキスト形式の入力テキストの選択開始位置を取得します。
    ///
    /// # 戻り値
    /// 選択開始位置（0から始まるインデックス）
    ///
    pub fn text_selection_start(&self) -> Result<i32> {
        Ok(unsafe { self.control.TextSelectionStart() }?)
    }

    /// テキスト形式の入力テキストの選択開始位置を設定します。
    ///
    /// # 引数
    /// * `value` - 設定する選択開始位置（0から始まるインデックス）
    ///
    pub fn set_text_selection_start(&self, value: i32) -> Result<()> {
        Ok(unsafe { self.control.SetTextSelectionStart(value) }?)
    }

    /// テキスト形式の入力テキストの選択文字数を取得します。
    ///
    /// # 戻り値
    /// 選択されているテキストの文字数
    ///
    pub fn text_selection_length(&self) -> Result<i32> {
        Ok(unsafe { self.control.TextSelectionLength() }?)
    }

    /// テキスト形式の入力テキストの選択文字数を設定します。
    ///
    /// # 引数
    /// * `value` - 設定する選択文字数
    ///
    pub fn set_text_selection_length(&self, value: i32) -> Result<()> {
        Ok(unsafe { self.control.SetTextSelectionLength(value) }?)
    }

    /// 現在のテキスト編集モードを取得します。
    ///
    /// # 戻り値
    /// `TextEditMode` 列挙型で表されるテキスト編集モード
    ///
    pub fn text_edit_mode(&self) -> Result<TextEditMode> {
        let text_edit_mode = unsafe { self.control.TextEditMode() }?;

        match text_edit_mode {
            ai_voice_sys::TextEditMode(0) => Ok(TextEditMode::Text),
            ai_voice_sys::TextEditMode(1) => Ok(TextEditMode::List),
            _ => anyhow::bail!("Unknown text edit mode"),
        }
    }

    /// テキスト編集モードを設定します。
    ///
    /// # 引数
    /// * `mode` - 設定する `TextEditMode` 列挙型のテキスト編集モード
    ///
    pub fn set_text_edit_mode(&self, mode: TextEditMode) -> Result<()> {
        let text_edit_mode = match mode {
            TextEditMode::Text => ai_voice_sys::TextEditMode(0),
            TextEditMode::List => ai_voice_sys::TextEditMode(1),
        };

        Ok(unsafe { self.control.SetTextEditMode(text_edit_mode) }?)
    }

    /// 音声の再生を開始または一時停止します。
    ///
    /// # 注意
    /// - ホストプログラムで選択されているテキスト入力形式で再生が行われます。
    /// - 再生の開始時、このメソッドは再生を開始すると終了し、再生の完了を待ちません。
    /// - ホストプログラムでフレーズが編集状態の場合、編集内容は破棄されます。
    /// - ホストプログラムで単語が編集状態の場合、その編集内容は読み上げに反映されません。
    ///
    pub fn play(&self) -> Result<()> {
        Ok(unsafe { self.control.Play() }?)
    }

    /// 音声の再生を停止します。
    ///
    pub fn stop(&self) -> Result<()> {
        Ok(unsafe { self.control.Stop() }?)
    }

    /// テキストの読み上げ音声を指定されたファイルに保存します。
    ///
    /// # 引数
    /// * `path` - 出力先ファイルパス
    ///
    /// # 注意
    /// - ホストプログラムで選択されているテキスト入力形式で保存が行われます。
    /// - 指定したファイルパスの拡張子がホストプログラムの音声保存時のファイル形式と一致しない場合、ファイル形式に応じた拡張子が付加されます。
    /// - ホストプログラムの「音声ファイルパスの指定方法」が「ファイル命名規則」の場合、引数で指定されたパスは無視されます。
    /// - ホストプログラムでフレーズや単語が編集状態の場合、その編集内容は読み上げに反映されません。
    ///
    pub fn save_audio_to_file(&self, path: &str) -> Result<()> {
        Ok(unsafe { self.control.SaveAudioToFile(&BSTR::from(path)) }?)
    }

    /// 読み上げ音声の再生時間を取得します。
    ///
    /// # 戻り値
    /// 再生時間（ミリ秒）
    ///
    /// # 注意
    /// - ホストプログラムで選択されているテキスト入力形式の再生時間を取得します。
    /// - ホストプログラムでフレーズや単語が編集状態の場合、その編集内容は
    ///   再生時間に反映されません。
    ///
    pub fn play_time(&self) -> Result<i64> {
        Ok(unsafe { self.control.GetPlayTime() }?)
    }

    /// リスト形式の行数を取得します。
    ///
    /// # 戻り値
    /// リスト形式の行数
    ///
    pub fn list_count(&self) -> Result<i32> {
        Ok(unsafe { self.control.GetListCount() }?)
    }

    /// リスト形式で選択されている行のインデックスを取得します。
    ///
    /// # 戻り値
    /// 選択行のインデックスのベクター（0スタート）
    ///
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

    /// リスト形式の選択行数を取得します。
    ///
    /// # 戻り値
    /// リスト形式の選択行数
    ///
    pub fn list_selection_count(&self) -> Result<i32> {
        Ok(unsafe { self.control.GetListSelectionCount() }?)
    }

    /// リスト形式の単一行を選択状態にします。
    ///
    /// # 引数
    /// * `index` - 選択状態にする行のインデックス（0スタート）
    ///
    /// # 注意
    /// 存在しないインデックスの指定は無視されます。
    ///
    pub fn set_list_selection_index(&self, index: i32) -> Result<()> {
        Ok(unsafe { self.control.SetListSelectionIndex(index) }?)
    }

    /// リスト形式の任意の複数行を選択状態にします。
    ///
    /// # 引数
    /// * `indices` - 選択状態にする行のインデックスのベクター（0スタート）
    ///
    /// # 注意
    /// 存在しないインデックスの指定は無視されます。
    ///
    pub fn set_list_selection_indices(&self, indices: Vec<String>) -> Result<()> {
        let rgsabound = vec![SAFEARRAYBOUND {
            cElements: indices.len() as u32,
            lLbound: 0,
        }];
        let psa = unsafe { SafeArrayCreate(VT_BSTR, 0, rgsabound.as_ptr()) };

        indices
            .iter()
            .enumerate()
            .map(|(i, elem)| unsafe {
                SafeArrayPutElement(
                    psa,
                    &(i as i32),
                    std::mem::transmute_copy(&BSTR::from(elem)),
                )
            })
            .collect::<Result<Vec<_>, _>>()?;

        Ok(unsafe { self.control.SetListSelectionIndices(psa) }?)
    }

    /// リスト形式の任意の範囲行を選択状態にします。
    ///
    /// # 引数
    /// * `startindex` - 選択開始行のインデックス（0スタート）
    /// * `length` - 選択状態にする行数
    ///
    /// # 注意
    /// 存在しないインデックスの指定は無視されます。
    ///
    pub fn set_list_selection_range(&self, startindex: i32, length: i32) -> Result<()> {
        Ok(unsafe { self.control.SetListSelectionRange(startindex, length) }?)
    }

    /// リスト形式の末尾に行を追加します。
    ///
    /// # 引数
    /// * `voice_preset_name` - ボイスプリセット名
    /// * `text` - テキスト
    ///
    pub fn add_list_item(&self, voice_preset_name: &str, text: &str) -> Result<()> {
        Ok(unsafe {
            self.control
                .AddListItem(&BSTR::from(voice_preset_name), &BSTR::from(text))
        }?)
    }

    /// リスト形式の選択位置に行を挿入します。
    ///
    /// # 引数
    /// * `voice_preset_name` - ボイスプリセット名
    /// * `text` - テキスト
    ///
    /// # 注意
    /// 単一行が選択されている場合のみ実行可能です。
    ///
    pub fn insert_list_item(&self, voice_preset_name: &str, text: &str) -> Result<()> {
        Ok(unsafe {
            self.control
                .InsertListItem(&BSTR::from(voice_preset_name), &BSTR::from(text))
        }?)
    }

    /// リスト形式の選択行を削除します。
    ///
    /// # 注意
    /// 選択された複数行に対して実行可能です。
    ///
    pub fn remove_list_item(&self) -> Result<()> {
        Ok(unsafe { self.control.RemoveListItem() }?)
    }

    /// リスト形式の行をすべて削除します。
    ///
    pub fn clear_list_items(&self) -> Result<()> {
        Ok(unsafe { self.control.ClearListItems() }?)
    }

    /// リスト形式の選択行のボイスプリセット名を取得します。
    ///
    /// # 戻り値
    /// ボイスプリセット名
    ///
    /// # 注意
    /// 単一行が選択されている場合のみ実行可能です。
    ///
    pub fn list_voice_preset(&self) -> Result<String> {
        Ok(unsafe { self.control.GetListVoicePreset() }?.to_string())
    }

    /// リスト形式の選択行のボイスプリセット名を設定します。
    ///
    /// # 引数
    /// * `voice_preset_name` - ボイスプリセット名
    ///
    /// # 注意
    /// 単一行が選択されている場合のみ実行可能です。
    ///
    pub fn set_list_voice_preset(&self, voice_preset_name: &str) -> Result<()> {
        Ok(unsafe {
            self.control
                .SetListVoicePreset(&BSTR::from(voice_preset_name))
        }?)
    }

    /// リスト形式の選択行のセンテンスを取得します。
    ///
    /// # 戻り値
    /// センテンス
    ///
    /// # 注意
    /// 単一行が選択されている場合のみ実行可能です。
    ///
    pub fn list_sentence(&self) -> Result<String> {
        Ok(unsafe { self.control.GetListSentence() }?.to_string())
    }

    /// 利用可能なボイス名を取得します。
    ///
    /// # 戻り値
    /// ボイス名のベクター
    ///
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

    /// 登録されているボイスプリセット名を取得します。
    ///
    /// # 戻り値
    /// ボイスプリセット名のベクター
    ///
    /// # 注意
    /// 標準ボイスプリセットとユーザーボイスプリセットの両方が含まれます。
    ///
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

    /// 現在のボイスプリセット名を取得します。
    ///
    /// # 戻り値
    /// 現在のボイスプリセット名
    ///
    pub fn current_voice_preset_name(&self) -> Result<String> {
        Ok(unsafe { self.control.CurrentVoicePresetName() }?.to_string())
    }

    /// 現在のボイスプリセット名を設定します。
    ///
    /// # 引数
    /// * `preset_name` - 設定するボイスプリセット名
    ///
    pub fn set_current_voice_preset_name(&self, preset_name: &str) -> Result<()> {
        Ok(unsafe {
            self.control
                .SetCurrentVoicePresetName(&BSTR::from(preset_name))
        }?)
    }

    /// 指定されたボイスプリセットの情報を取得します。
    ///
    /// # 引数
    /// * `preset_name` - 取得するボイスプリセットの名前
    ///
    /// # 戻り値
    /// `VoicePreset`構造体で表されるボイスプリセットの情報
    ///
    /// # エラー
    /// ボイスプリセットの解析に失敗した場合にエラーを返します。
    ///
    pub fn voice_preset(&self, preset_name: &str) -> Result<VoicePreset> {
        let voice_preset =
            unsafe { self.control.GetVoicePreset(&BSTR::from(preset_name)) }?.to_string();
        serde_json::from_str(&voice_preset).with_context(|| "Failed to parse voice preset")
    }
    /// 既存のボイスプリセットに指定された設定を適用します。
    ///
    /// # 引数
    /// * `voice_preset` - 適用する`VoicePreset`構造体
    pub fn set_voice_preset(&self, voice_preset: &VoicePreset) -> Result<()> {
        let json = serde_json::to_string(voice_preset)?;
        Ok(unsafe { self.control.SetVoicePreset(&BSTR::from(json)) }?)
    }

    /// 新規ボイスプリセットを作成します。
    ///
    /// # 引数
    /// * `voice_preset` - 作成する`VoicePreset`構造体
    ///
    pub fn add_voice_preset(&self, voice_preset: &VoicePreset) -> Result<()> {
        let json = serde_json::to_string(voice_preset)?;
        Ok(unsafe { self.control.AddVoicePreset(&BSTR::from(json)) }?)
    }

    /// ボイスプリセットを再読込みします。
    ///
    pub fn reload_voice_presets(&self) -> Result<()> {
        Ok(unsafe { self.control.ReloadVoicePresets() }?)
    }

    /// フレーズ辞書を再読込みします。
    ///
    pub fn reload_phrase_dictionary(&self) -> Result<()> {
        Ok(unsafe { self.control.ReloadPhraseDictionary() }?)
    }

    /// 単語辞書を再読込みします。
    ///
    pub fn reload_word_dictionary(&self) -> Result<()> {
        Ok(unsafe { self.control.ReloadWordDictionary() }?)
    }

    /// 記号ポーズ辞書を再読込みします。
    ///
    pub fn reload_symbol_dictionary(&self) -> Result<()> {
        Ok(unsafe { self.control.ReloadSymbolDictionary() }?)
    }
}
