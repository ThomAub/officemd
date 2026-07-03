class OfficemdCli < Formula
  desc "CLI for OfficeMD document extraction and markdown rendering"
  homepage "https://github.com/ThomAub/officemd"
  version "0.1.8"
  if OS.mac?
    if Hardware::CPU.arm?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.8/officemd_cli-aarch64-apple-darwin.tar.xz"
      sha256 "d02d9889c5568de6a8038754ec4695880e81ac6f2862914b1f6d1898b68edb58"
    end
    if Hardware::CPU.intel?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.8/officemd_cli-x86_64-apple-darwin.tar.xz"
      sha256 "f01a24fd2fda2f565201ec5ee3337fafdc296d72691855df9b689ef1570a6145"
    end
  end
  if OS.linux?
    if Hardware::CPU.arm?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.8/officemd_cli-aarch64-unknown-linux-gnu.tar.xz"
      sha256 "4bf74e3d7b7be9faa0dbfaa07a48285813c37c2fe8bedec6afb5926e87654925"
    end
    if Hardware::CPU.intel?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.8/officemd_cli-x86_64-unknown-linux-gnu.tar.xz"
      sha256 "cfea65ea16580a10fa29d82dfc41a660cb422fdfed16ee89b9344f40e1db391e"
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
