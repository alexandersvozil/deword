class Deword < Formula
  desc "🪱 De-Words your documents for AI agents — create, read, edit, fill, and add equations to Word docs"
  homepage "https://github.com/alexandersvozil/deword"
  url "https://registry.npmjs.org/deword/-/deword-0.3.1.tgz"
  sha256 "dc13cfe8ceb80474bfa1a98e6ccf61ea18874c474b92fcefdb5c113887d270af"
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
