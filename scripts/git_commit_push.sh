#!/bin/zsh
# Simple script to add, commit, and push changes in one go

# Change to the repository root directory
cd "$(git rev-parse --show-toplevel)" || exit 1

# Add all changes
echo "📦 Adding all changes..."
git add .

# Ask for commit message
echo "💬 Enter your commit message:"
read -r commit_message

# Check if commit message is provided
if [[ -z "$commit_message" ]]; then
  echo "❌ Error: Commit message cannot be empty"
  exit 1
fi

# Commit with the provided message
echo "✅ Committing changes..."
git commit -m "$commit_message"

# Push to the current branch
current_branch=$(git symbolic-ref --short HEAD)
echo "🚀 Pushing to branch: $current_branch..."
git push origin "$current_branch"

echo "✨ Done! All changes have been added, committed, and pushed."
