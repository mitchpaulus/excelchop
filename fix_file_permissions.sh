#!/bin/sh

# An unfortunate side effect of working in the WSL is that the file permissions can get jacked.
# This script will fix that.

fd -e cs -e csproj -e sln -x chmod 644
