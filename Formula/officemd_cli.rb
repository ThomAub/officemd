class OfficemdCli < Formula
  desc "CLI for OfficeMD document extraction and markdown rendering"
  homepage "https://github.com/ThomAub/officemd"
  version "0.1.6"
  if OS.mac?
    if Hardware::CPU.arm?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.6/officemd_cli-aarch64-apple-darwin.tar.xz"
      sha256 "2bdefced81a24d769db69db4dca3dc9384a4e6c7b67bcfb3be9e57bcbbddbf69"
    end
    if Hardware::CPU.intel?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.6/officemd_cli-x86_64-apple-darwin.tar.xz"
      sha256 "46db9e4f6de4360c7ba6fce3800abcc8bdfb116d0652e3d9ad7f5fd9dcf37b3f"
    end
  end
  if OS.linux?
    if Hardware::CPU.arm?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.6/officemd_cli-aarch64-unknown-linux-gnu.tar.xz"
      sha256 "b4d3bcd9f45123394d95913ba52d800f0eea19772d3ebca58cefef2afbbc316c"
    end
    if Hardware::CPU.intel?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.6/officemd_cli-x86_64-unknown-linux-gnu.tar.xz"
      sha256 "0acfe10cfd375e21b5aa88ffe9269f199aa8148c64470ce2f4e220695d440909"
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
