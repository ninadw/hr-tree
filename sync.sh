#!/usr/bin/env bash
# 一键同步：把本地技能目录同步到本仓库并推送到 GitHub
# 用法：./sync.sh ["提交说明"]      （不填说明则自动生成）
set -euo pipefail

SKILL_DIR="${SKILL_DIR:-$HOME/.comate/skills/architecture-tree}"
REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
MSG="${1:-}"

if [ ! -d "$SKILL_DIR" ]; then
  echo "✗ 找不到技能目录：$SKILL_DIR"
  echo "  可用 SKILL_DIR=/path/to/skill ./sync.sh 指定"
  exit 1
fi

cd "$REPO_DIR"

echo "→ 同步文件：$SKILL_DIR → $REPO_DIR"
mkdir -p assets scripts references
cp "$SKILL_DIR/SKILL.md" ./SKILL.md
cp "$SKILL_DIR/assets/org-chart-template.html" ./assets/
cp "$SKILL_DIR/scripts/bake_data.py" ./scripts/
[ -f "$SKILL_DIR/scripts/new_page.py" ] && cp "$SKILL_DIR/scripts/new_page.py" ./scripts/
cp "$SKILL_DIR/references/data-format.md" ./references/
find . -name '__pycache__' -type d -prune -exec rm -rf {} + 2>/dev/null || true
find . -name '*.pyc' -delete 2>/dev/null || true

if git diff --quiet && git diff --cached --quiet && [ -z "$(git status --porcelain)" ]; then
  echo "✓ 没有变化，无需提交"
else
  git add -A
  if [ -z "$MSG" ]; then
    CHANGED=$(git diff --cached --name-only | tr '\n' ' ')
    MSG="chore: 同步技能更新（$(date '+%Y-%m-%d %H:%M')）"
    echo "  变更文件：$CHANGED"
  fi
  git commit -q -m "$MSG"
  echo "✓ 已提交：$MSG"
fi

echo "→ 同步远程并推送…"
git fetch -q origin
if ! git rebase origin/main >/dev/null 2>&1; then
  echo "✗ rebase 失败（可能有冲突）。请手动解决后执行：git rebase --continue && git push"
  git rebase --abort 2>/dev/null || true
  exit 1
fi

if git push origin main 2>&1 | tail -3; then
  echo "✓ 已推送到 GitHub：$(git config --get remote.origin.url)"
  git log --oneline -1
else
  echo "✗ 推送失败：请检查网络与 GitHub 凭据（HTTPS 需 Personal Access Token）"
  exit 1
fi
