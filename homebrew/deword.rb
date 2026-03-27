class Deword < Formula
  desc "🪱 De-Words your documents for AI agents — read, edit, fill Word docs"
  homepage "https://github.com/alexandersvozil/deword"
  url "https://registry.npmjs.org/deword/-/deword-0.2.1.tgz"
  sha256 "5090fa8d081b64b74af5f5575838a99c8e2dcaa8fde0fc445e0434c2a323e58a"
  license "MIT"

  depends_on "node"

  def install
    system "npm", "install", *std_npm_args
    bin.install_symlink Dir["#{libexec}/bin/*"]
  end

  test do
    assert_match "deword", shell_output("#{bin}/deword --version")
  end
end
