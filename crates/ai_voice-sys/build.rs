fn main() {
    println!("cargo:rerun-if-changed=.windows/winmd/AI.Talk.Editor.Api.winmd");
    println!("cargo:rerun-if-changed=build.rs");

    windows_bindgen::bindgen([
        "--in",
        ".windows/winmd/AI.Talk.Editor.Api.winmd",
        "--out",
        "src/bindings.rs",
        "--filter",
        "AI.Talk.Editor.Api",
        "--config",
        "implement",
        "--config",
        "vtbl",
    ])
    .unwrap();
}
