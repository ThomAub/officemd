fn main() {
    if std::env::var_os("CARGO_FEATURE_EXTENSION_MODULE").is_some() {
        // Required on macOS for PyO3 extension modules to avoid unresolved _Py* symbols.
        pyo3_build_config::add_extension_module_link_args();
    } else if std::env::var_os("CARGO_CFG_TARGET_OS").as_deref()
        == Some(std::ffi::OsStr::new("macos"))
    {
        // Rust test binaries link against libpython and need a runtime search path.
        if let Some(lib_dir) = pyo3_build_config::get().lib_dir.as_deref() {
            println!("cargo:rustc-link-arg=-Wl,-rpath,{lib_dir}");
        }

        // Handle Python.framework-based installs.
        pyo3_build_config::add_python_framework_link_args();
    }
}
