/// Check if AI mode is active via the `AI` environment variable.
///
/// Returns `true` if `AI=True` (case-insensitive) or `AI=1`.
///
/// # Examples
///
/// ```no_run
/// if clap_ai::is_ai_mode() {
///     eprintln!("Running in AI mode");
/// }
/// ```
#[must_use]
pub fn is_ai_mode() -> bool {
    std::env::var("AI")
        .map(|v| v.eq_ignore_ascii_case("true") || v == "1")
        .unwrap_or(false)
}

/// Trait for types that can have AI-friendly defaults applied.
///
/// Implement this on your CLI options struct to define what "AI mode" means
/// for your application. Typically this means switching to structured output
/// (JSON) and enabling pretty-printing.
///
/// # Examples
///
/// ```
/// use clap_ai::AiDefaults;
///
/// struct MyOptions {
///     json: bool,
///     pretty: bool,
/// }
///
/// impl AiDefaults for MyOptions {
///     fn apply_ai_defaults(&mut self) {
///         self.json = true;
///         self.pretty = true;
///     }
/// }
/// ```
pub trait AiDefaults {
    fn apply_ai_defaults(&mut self);
}

/// Check AI mode and apply defaults if active.
///
/// Returns whether AI mode was active.
pub fn maybe_apply_ai_defaults<T: AiDefaults>(args: &mut T) -> bool {
    let active = is_ai_mode();
    if active {
        args.apply_ai_defaults();
    }
    active
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ai_mode_trait_applies_defaults() {
        struct Opts {
            json: bool,
        }
        impl AiDefaults for Opts {
            fn apply_ai_defaults(&mut self) {
                self.json = true;
            }
        }
        let mut opts = Opts { json: false };
        opts.apply_ai_defaults();
        assert!(opts.json);
    }
}
