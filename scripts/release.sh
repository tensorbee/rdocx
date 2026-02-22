#!/usr/bin/env bash
set -euo pipefail

# Usage: ./scripts/release.sh <version>
# Example: ./scripts/release.sh 0.2.0
#
# Bumps the version in all Cargo.toml files, commits, tags, and pushes.

if [ $# -ne 1 ]; then
    echo "Usage: $0 <version>"
    echo "Example: $0 0.2.0"
    exit 1
fi

VERSION="$1"

# Validate semver format
if ! [[ "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "Error: version must be semver (e.g. 0.2.0), got: $VERSION"
    exit 1
fi

# Ensure clean working tree
if [ -n "$(git status --porcelain)" ]; then
    echo "Error: working tree is not clean. Commit or stash changes first."
    exit 1
fi

# Ensure on main branch
BRANCH=$(git rev-parse --abbrev-ref HEAD)
if [ "$BRANCH" != "main" ]; then
    echo "Error: not on main branch (currently on $BRANCH)"
    exit 1
fi

# Check tag doesn't already exist
if git rev-parse "v$VERSION" >/dev/null 2>&1; then
    echo "Error: tag v$VERSION already exists"
    exit 1
fi

echo "Bumping version to $VERSION..."

# Get current version from workspace Cargo.toml
CURRENT=$(grep '^version = ' Cargo.toml | head -1 | sed 's/version = "\(.*\)"/\1/')
echo "Current version: $CURRENT"

# Update workspace Cargo.toml (package version + internal dep versions)
sed -i '' "s/version = \"$CURRENT\"/version = \"$VERSION\"/g" Cargo.toml

# Update rdocx-wasm (excluded from workspace, has its own version)
sed -i '' "s/version = \"$CURRENT\"/version = \"$VERSION\"/" crates/rdocx-wasm/Cargo.toml

# Regenerate lockfile
cargo check --workspace --quiet

echo "Updated all crates to $VERSION"

# Commit, tag, and push
git add Cargo.toml Cargo.lock crates/rdocx-wasm/Cargo.toml
git commit -m "Release v$VERSION"
git tag "v$VERSION"
git push origin main "v$VERSION"

echo ""
echo "Done! v$VERSION pushed. The publish workflow will run automatically."
echo "Monitor at: https://github.com/tensorbee/rdocx/actions"
