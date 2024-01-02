#!/bin/sh

set -e
mkdir -p ~/.local/bin
curl -sL -o /tmp/excelchop.zip 'https://github.com/mitchpaulus/excelchop/releases/download/v0.8.0/excelchop-linux-x64-v0.8.0-framework-dependent.zip'
unzip -p /tmp/excelchop.zip excelchop > ~/.local/bin/excelchop
chmod 755 ~/.local/bin/excelchop
rm /tmp/excelchop.zip
