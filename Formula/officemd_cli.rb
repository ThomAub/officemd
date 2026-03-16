class OfficemdCli < Formula
  desc "CLI for OfficeMD document extraction and markdown rendering"
  homepage "https://github.com/ThomAub/officemd"
  version "0.1.4"
  if OS.mac?
    if Hardware::CPU.arm?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.4/officemd_cli-aarch64-apple-darwin.tar.xz"
      sha256 "16d340bececc75fac321704a33615e39a68dc790117534f869c0552990a38d1f"
    end
    if Hardware::CPU.intel?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.4/officemd_cli-x86_64-apple-darwin.tar.xz"
      sha256 "f31e55e37e65af4d5f3cd671b4be5d664a5e89c51c09797fe3fdae31e22ae0db"
    end
  end
  if OS.linux?
    if Hardware::CPU.arm?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.4/officemd_cli-aarch64-unknown-linux-gnu.tar.xz"
      sha256 "46d991e657360471ac8071c8e6ec3e017655e24c670b416d2eabf5e554c3d907"
    end
    if Hardware::CPU.intel?
      url "https://github.com/ThomAub/officemd/releases/download/v0.1.4/officemd_cli-x86_64-unknown-linux-gnu.tar.xz"
      sha256 "5ee20b0f165215b0bae5c4363ea655ad7f78e4e4ffe35c5678c25cc392cf9f28"
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
