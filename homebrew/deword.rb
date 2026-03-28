class Deword < Formula
  desc "🪱 De-Words your documents for AI agents — read, edit, fill Word docs"
  homepage "https://github.com/alexandersvozil/deword"
  url "https://registry.npmjs.org/deword/-/deword-0.3.0.tgz"
  sha256 "0f700fccd045c38733b6d3357f469c7a05db75a9545c5351cd7e5e7a28fd76e2"
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
