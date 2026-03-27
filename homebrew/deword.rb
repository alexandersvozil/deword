class Deword < Formula
  desc "🪱 De-Words your documents for AI agents — Word to markdown CLI"
  homepage "https://github.com/alexandersvozil/deword"
  url "https://registry.npmjs.org/deword/-/deword-0.1.0.tgz"
  sha256 "422f9812ff76f99ad8eabebc646121a2592d5b017f0f7869f99f75dd6b8485d9"
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
