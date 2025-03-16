fn main() {
    println!("cargo:rerun-if-changed=.windows/winmd/AI.Talk.Editor.Api.winmd");
    println!("cargo:rerun-if-changed=build.rs");

    windows_bindgen::bindgen([
        "--in",
        "default",
        ".windows/winmd/AI.Talk.Editor.Api.winmd",
        "--out",
        "src/bindings.rs",
        "--filter",
        "AI.Talk.Editor.Api",
        "--implement",
        "--flat",
        "--reference",
        "windows,skip-root,Windows.Win32.System.Com",
        "--reference",
        "windows,skip-root,Windows.Win32.Foundation",
    ]);
}
