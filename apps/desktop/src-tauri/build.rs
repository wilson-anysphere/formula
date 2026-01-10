fn main() {
    // Only needed when building the desktop binary target. Keeping it in place
    // matches the standard Tauri layout.
    #[cfg(feature = "desktop")]
    tauri_build::build();
}
