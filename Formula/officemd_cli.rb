class OfficemdCli < Formula
  desc "CLI for OfficeMD document extraction and markdown rendering"
  homepage "https://github.com/ThomAub/officemd"
  version "0.1.5"
  if OS.mac?
    if Hardware::CPU.arm?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.5/officemd_cli-aarch64-apple-darwin.tar.xz"
      sha256 "e8e6d85a51d6fadda95a203ee02575a3616af931648f5599f33a47cf266379b2"
    end
    if Hardware::CPU.intel?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.5/officemd_cli-x86_64-apple-darwin.tar.xz"
      sha256 "9aa6cbc06f82cd5adaa1643a5b5583d86c5f6accf7ab630d680532b071b66063"
    end
  end
  if OS.linux?
    if Hardware::CPU.arm?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.5/officemd_cli-aarch64-unknown-linux-gnu.tar.xz"
      sha256 "d9fbbc4cca7fa3538d4964e14cd24462144e34c4213cf97254da753867b91713"
    end
    if Hardware::CPU.intel?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.5/officemd_cli-x86_64-unknown-linux-gnu.tar.xz"
      sha256 "8591c7be45b3ac17de03cc658f2ae52e696f88c9a1d5317d2c05f5b7f63ee2ea"
    end
  end
  license "MIT"

  BINARY_ALIASES = {
    "aarch64-apple-darwin": {},
    "aarch64-pc-windows-gnu": {},
    "aarch64-unknown-linux-gnu": {},
    "aarch64-unknown-linux-musl-dynamic": {},
    "aarch64-unknown-linux-musl-static": {},
    "x86_64-apple-darwin": {},
    "x86_64-pc-windows-gnu": {},
    "x86_64-unknown-linux-gnu": {},
    "x86_64-unknown-linux-musl-dynamic": {},
    "x86_64-unknown-linux-musl-static": {}
  }

  def target_triple
    cpu = Hardware::CPU.arm? ? "aarch64" : "x86_64"
    os = OS.mac? ? "apple-darwin" : "unknown-linux-gnu"

    "#{cpu}-#{os}"
  end

  def install_binary_aliases!
    BINARY_ALIASES[target_triple.to_sym].each do |source, dests|
      dests.each do |dest|
        bin.install_symlink bin/source.to_s => dest
      end
    end
  end

  def install
    if OS.mac? && Hardware::CPU.arm?
      bin.install "officemd"
    end
    if OS.mac? && Hardware::CPU.intel?
      bin.install "officemd"
    end
    if OS.linux? && Hardware::CPU.arm?
      bin.install "officemd"
    end
    if OS.linux? && Hardware::CPU.intel?
      bin.install "officemd"
    end

    install_binary_aliases!

    # Homebrew will automatically install these, so we don't need to do that
    doc_files = Dir["README.*", "readme.*", "LICENSE", "LICENSE.*", "CHANGELOG.*"]
    leftover_contents = Dir["*"] - doc_files

    # Install any leftover files in pkgshare; these are probably config or
    # sample files.
    pkgshare.install(*leftover_contents) unless leftover_contents.empty?
  end
end
