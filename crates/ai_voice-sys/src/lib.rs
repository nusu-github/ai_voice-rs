use windows::core::GUID;

pub use bindings::Talk::Editor::Api::{
    ITtsControl, __MIDL___MIDL_itf_AI2ETalk2EEditor2EApi_0000_0000_0001 as HostStatus,
    __MIDL___MIDL_itf_AI2ETalk2EEditor2EApi_0000_0000_0002 as TextEditMode,
};

mod bindings;

pub const CLSID_TTS_CONTROL: GUID = GUID::from_u128(0xB628D293_341C_41BE_B2E7_9E7822B2B7AC);
