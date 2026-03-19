#!/usr/bin/env node

import { mkdirSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";

const [, , rawVersion, darwinX64Sha256, darwinArm64Sha256, outputPath] = process.argv;

if (!rawVersion || !darwinX64Sha256 || !darwinArm64Sha256) {
  console.error(
    "Usage: node scripts/generate-homebrew-formula.mjs <version> <darwin_x64_sha256> <darwin_arm64_sha256> [output_path]"
  );
  process.exit(1);
}

const version = rawVersion.startsWith("v") ? rawVersion : `v${rawVersion}`;
const releaseBaseUrl = `https://github.com/t-izuno/owa-mail/releases/download/${version}`;
const darwinX64Archive = `owa-mail_${version}_darwin_x64.tar.gz`;
const darwinArm64Archive = `owa-mail_${version}_darwin_arm64.tar.gz`;
const output = outputPath ?? "owa-mail.rb";

const formula = `class OwaMail < Formula
  desc "Office365 OWA mail client using Playwright session"
  homepage "https://github.com/t-izuno/owa-mail"
  license "MIT"

  on_macos do
    if Hardware::CPU.arm?
      url "${releaseBaseUrl}/${darwinArm64Archive}"
      sha256 "${darwinArm64Sha256}"
    else
      url "${releaseBaseUrl}/${darwinX64Archive}"
      sha256 "${darwinX64Sha256}"
    end
  end

  def install
    bin.install "owa-mail"
  end

  def caveats
    <<~EOS
      Login requires Playwright CLI.
      Install it separately with:
        npm install -g @playwright/cli@latest
    EOS
  end

  test do
    assert_match "owa-mail <command> [options]", shell_output("#{bin}/owa-mail help")
  end
end
`;

mkdirSync(dirname(output), { recursive: true });
writeFileSync(output, formula);
