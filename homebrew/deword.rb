class Deword < Formula
  desc "🪱 De-Words your documents for AI agents — read, edit, fill Word docs"
  homepage "https://github.com/alexandersvozil/deword"
  url "https://registry.npmjs.org/deword/-/deword-0.2.0.tgz"
  sha256 "8753764d06c1aae36fe0d4f99aac729d054452a2f46047634abd8ac7c32fa98b"
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
