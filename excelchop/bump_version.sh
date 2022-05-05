#!/bin/sh

# Argument 1 is new version num. (e.g. 1.0.0)

VERSION="$1"

# Verify the version number is valid
if [ -z "$VERSION" ]; then
  echo "Usage: $0 <version>"
  exit 1
fi

if printf "%s" "$VERSION" | grep -qvE '^[0-9]+\.[0-9]+\.[0-9]+$'; then
  echo "Invalid version number: $VERSION"
  exit 1
fi

CURRENT_ISO_DATE=$(date -u +%Y-%m-%d)

sed -i -e "s/[0-9]\+\.[0-9]\+\.[0-9]\+ - [0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]/$VERSION - $CURRENT_ISO_DATE/g" Program.cs
