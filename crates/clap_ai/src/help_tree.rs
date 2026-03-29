use clap::Command;
use std::fmt::Write;

/// Options controlling help tree rendering.
#[derive(Debug, Clone)]
pub struct HelpTreeOptions {
    /// Maximum depth to render. 1 = command names only, 2 = include args/options.
    /// Depths beyond 2 recurse into nested subcommands.
    pub depth: u8,
    /// Optional suffix appended to the root line (e.g., `" [AI mode]"`).
    pub root_suffix: Option<String>,
    /// Optional footer lines printed after the tree.
    pub footer_lines: Vec<String>,
}

impl Default for HelpTreeOptions {
    fn default() -> Self {
        Self {
            depth: 2,
            root_suffix: None,
            footer_lines: Vec::new(),
        }
    }
}

/// Render a help tree from a clap [`Command`] and print it to stdout.
pub fn print_help_tree(cmd: &Command, opts: &HelpTreeOptions) {
    print!("{}", render_help_tree(cmd, opts));
}

/// Render a help tree from a clap [`Command`] into a [`String`].
///
/// Works with both derive and builder pattern CLIs — only requires a `&Command`.
///
/// # Examples
///
/// **Derive pattern:**
/// ```no_run
/// use clap::{Parser, CommandFactory};
///
/// #[derive(Parser)]
/// #[command(about = "My CLI")]
/// struct Cli {
///     #[command(subcommand)]
///     command: Option<MyCommand>,
/// }
///
/// #[derive(clap::Subcommand)]
/// enum MyCommand {
///     /// Do something
///     Foo { name: String },
/// }
///
/// let cmd = Cli::command();
/// let tree = clap_ai::render_help_tree(&cmd, &clap_ai::HelpTreeOptions::default());
/// ```
///
/// **Builder pattern:**
/// ```
/// use clap::{Command, Arg};
///
/// let cmd = Command::new("myapp")
///     .about("My CLI")
///     .subcommand(Command::new("foo").about("Do something"));
///
/// let tree = clap_ai::render_help_tree(&cmd, &clap_ai::HelpTreeOptions::default());
/// assert!(tree.contains("foo"));
/// ```
#[must_use]
pub fn render_help_tree(cmd: &Command, opts: &HelpTreeOptions) -> String {
    let mut out = String::new();

    // Root line
    let name = cmd.get_name();
    let suffix = opts.root_suffix.as_deref().unwrap_or("");
    if let Some(about) = cmd.get_about() {
        let _ = writeln!(out, "{name} \u{2014} {about}{suffix}");
    } else {
        let _ = writeln!(out, "{name}{suffix}");
    }
    let _ = writeln!(out);

    // Subcommands
    let subs: Vec<&Command> = cmd.get_subcommands().filter(|s| !s.is_hide_set()).collect();

    if subs.is_empty() {
        // No subcommands — render top-level args directly
        render_args(&mut out, cmd, "");
    } else {
        let last_idx = subs.len() - 1;
        for (i, sub) in subs.iter().enumerate() {
            let is_last = i == last_idx;
            render_subcommand(&mut out, sub, is_last, opts.depth, "", 1);
        }
    }

    // Root-level flags from root command, split into global and non-global
    let root_args: Vec<_> = cmd
        .get_arguments()
        .filter(|a| !a.is_hide_set() && !a.is_positional())
        .collect();
    let global_args: Vec<_> = root_args.iter().filter(|a| a.is_global_set()).collect();
    let local_args: Vec<_> = root_args.iter().filter(|a| !a.is_global_set()).collect();
    if !global_args.is_empty() {
        let _ = writeln!(out);
        let _ = writeln!(out, "Global flags:");
        for arg in global_args {
            let _ = writeln!(out, "  {}", format_flag(arg));
        }
    }
    if !local_args.is_empty() {
        let _ = writeln!(out);
        let _ = writeln!(out, "Root options:");
        for arg in local_args {
            let _ = writeln!(out, "  {}", format_flag(arg));
        }
    }

    // Footer
    for line in &opts.footer_lines {
        let _ = writeln!(out, "{line}");
    }

    out
}

fn render_subcommand(
    out: &mut String,
    cmd: &Command,
    is_last: bool,
    max_depth: u8,
    parent_prefix: &str,
    current_depth: u8,
) {
    let (branch, cont) = if is_last {
        ("\u{2514}\u{2500}\u{2500}", "    ")
    } else {
        ("\u{251c}\u{2500}\u{2500}", "\u{2502}   ")
    };

    let child_prefix = format!("{parent_prefix}{cont}");

    let _ = writeln!(out, "{parent_prefix}{branch} {}", cmd.get_name());

    if current_depth >= max_depth {
        return;
    }

    // About text
    if let Some(about) = cmd.get_about() {
        let _ = writeln!(out, "{child_prefix} {about}");
    }

    // Arguments
    render_args(out, cmd, &child_prefix);

    // Nested subcommands
    let nested: Vec<&Command> = cmd.get_subcommands().filter(|s| !s.is_hide_set()).collect();
    if !nested.is_empty() {
        let last_idx = nested.len() - 1;
        for (i, sub) in nested.iter().enumerate() {
            let nested_last = i == last_idx;
            render_subcommand(
                out,
                sub,
                nested_last,
                max_depth,
                &child_prefix,
                current_depth + 1,
            );
        }
    }

    // Blank line separator between sibling commands
    if !is_last {
        let _ = writeln!(out, "{parent_prefix}{cont}");
    }
}

fn render_args(out: &mut String, cmd: &Command, prefix: &str) {
    // Positional args first
    for arg in cmd
        .get_arguments()
        .filter(|a| a.is_positional() && !a.is_hide_set())
    {
        let _ = writeln!(out, "{prefix} {}", format_positional(arg));
    }

    // Then flags/options
    let flags: Vec<_> = cmd
        .get_arguments()
        .filter(|a| !a.is_positional() && !a.is_hide_set())
        .collect();
    if !flags.is_empty() {
        for flag in flags {
            let _ = writeln!(out, "{prefix}   {}", format_flag(flag));
        }
    }
}

fn format_positional(arg: &clap::Arg) -> String {
    let name = arg.get_value_names().and_then(|v| v.first()).map_or_else(
        || format!("<{}>", arg.get_id().as_str().to_uppercase()),
        |v| format!("<{v}>"),
    );
    let help = arg
        .get_help()
        .map(std::string::ToString::to_string)
        .unwrap_or_default();
    if help.is_empty() {
        name
    } else {
        format!("{name:30} {help}")
    }
}

fn format_flag(arg: &clap::Arg) -> String {
    let mut flag = String::new();

    if let Some(short) = arg.get_short() {
        let _ = write!(flag, "-{short}, ");
    }
    if let Some(long) = arg.get_long() {
        let _ = write!(flag, "--{long}");
    }
    // Show value names, but skip for boolean/count flags
    let is_bool_flag = matches!(
        arg.get_action(),
        clap::ArgAction::SetTrue | clap::ArgAction::SetFalse | clap::ArgAction::Count
    );
    if !is_bool_flag && let Some(vals) = arg.get_value_names() {
        for v in vals {
            let _ = write!(flag, " <{v}>");
        }
    }

    let help = arg
        .get_help()
        .map(std::string::ToString::to_string)
        .unwrap_or_default();
    if help.is_empty() {
        flag
    } else {
        format!("{flag:36} {help}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use clap::{Arg, Command};

    fn sample_cmd() -> Command {
        Command::new("myapp")
            .about("A sample CLI")
            .subcommand(
                Command::new("greet")
                    .about("Say hello")
                    .arg(Arg::new("name").required(true).help("Who to greet")),
            )
            .subcommand(
                Command::new("serve")
                    .about("Start a server")
                    .arg(
                        Arg::new("port")
                            .long("port")
                            .short('p')
                            .value_name("PORT")
                            .help("Port to listen on"),
                    )
                    .arg(
                        Arg::new("host")
                            .long("host")
                            .value_name("HOST")
                            .help("Host to bind"),
                    ),
            )
    }

    #[test]
    fn depth_1_shows_command_names_only() {
        let opts = HelpTreeOptions {
            depth: 1,
            ..Default::default()
        };
        let tree = render_help_tree(&sample_cmd(), &opts);
        assert!(tree.contains("\u{251c}\u{2500}\u{2500} greet"));
        assert!(tree.contains("\u{2514}\u{2500}\u{2500} serve"));
        // Should NOT contain argument details
        assert!(!tree.contains("Say hello"));
        assert!(!tree.contains("--port"));
    }

    #[test]
    fn depth_2_shows_args_and_help() {
        let opts = HelpTreeOptions::default();
        let tree = render_help_tree(&sample_cmd(), &opts);
        assert!(tree.contains("Say hello"));
        assert!(tree.contains("<NAME>"));
        assert!(tree.contains("Who to greet"));
        assert!(tree.contains("--port"));
        assert!(tree.contains("--host"));
        assert!(tree.contains("Start a server"));
    }

    #[test]
    fn root_suffix_appears() {
        let opts = HelpTreeOptions {
            root_suffix: Some(" [AI mode]".into()),
            ..Default::default()
        };
        let tree = render_help_tree(&sample_cmd(), &opts);
        assert!(tree.contains("A sample CLI [AI mode]"));
    }

    #[test]
    fn footer_lines_appear() {
        let opts = HelpTreeOptions {
            footer_lines: vec!["Custom footer".into()],
            ..Default::default()
        };
        let tree = render_help_tree(&sample_cmd(), &opts);
        assert!(tree.contains("Custom footer"));
    }

    #[test]
    fn nested_subcommands_rendered() {
        let cmd = Command::new("app").about("Nested CLI").subcommand(
            Command::new("group")
                .about("A group of commands")
                .subcommand(Command::new("sub1").about("First sub"))
                .subcommand(Command::new("sub2").about("Second sub")),
        );
        let opts = HelpTreeOptions {
            depth: 3,
            ..Default::default()
        };
        let tree = render_help_tree(&cmd, &opts);
        assert!(tree.contains("group"));
        assert!(tree.contains("sub1"));
        assert!(tree.contains("First sub"));
        assert!(tree.contains("sub2"));
    }

    #[test]
    fn global_vs_root_flags_labeled_correctly() {
        let cmd = Command::new("app")
            .about("Test")
            .subcommand(Command::new("sub").about("Sub"))
            .arg(
                Arg::new("verbose")
                    .long("verbose")
                    .global(true)
                    .action(clap::ArgAction::SetTrue)
                    .help("Verbose output"),
            )
            .arg(
                Arg::new("config")
                    .long("config")
                    .value_name("PATH")
                    .help("Config file"),
            );
        let tree = render_help_tree(&cmd, &HelpTreeOptions::default());
        assert!(
            tree.contains("Global flags:"),
            "should have Global flags section"
        );
        assert!(tree.contains("--verbose"), "global flag should appear");
        assert!(
            tree.contains("Root options:"),
            "should have Root options section"
        );
        assert!(tree.contains("--config"), "root option should appear");
    }

    #[test]
    fn hidden_commands_excluded() {
        let cmd = Command::new("app")
            .about("Test")
            .subcommand(Command::new("visible").about("Shown"))
            .subcommand(Command::new("secret").about("Hidden").hide(true));
        let tree = render_help_tree(&cmd, &HelpTreeOptions::default());
        assert!(tree.contains("visible"));
        assert!(!tree.contains("secret"));
    }
}
