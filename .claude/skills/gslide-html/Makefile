# Install the gslide-html skill into the user-global Claude Code skills dir
# (~/.claude/skills), so it is available in every project / folder on this
# machine. On a new device: clone the repo, then
#
#     make -C .claude/skills/gslide-html install
#
# Targets:
#   install       symlink into ~/.claude/skills (stays in sync with the repo)
#   install-copy  copy instead of symlink (survives moving/deleting the repo)
#   uninstall     remove the installed skill
#   status        show what is currently installed

SKILL_NAME := gslide-html
SKILL_DIR  := $(abspath $(dir $(lastword $(MAKEFILE_LIST))))
SKILLS_HOME := $(HOME)/.claude/skills
TARGET      := $(SKILLS_HOME)/$(SKILL_NAME)

.PHONY: install install-copy uninstall status

install:
	mkdir -p $(SKILLS_HOME)
	ln -sfn $(SKILL_DIR) $(TARGET)
	@echo "Installed (symlink): $(TARGET) -> $(SKILL_DIR)"
	@echo "Restart your Claude Code session to pick it up."

install-copy:
	mkdir -p $(SKILLS_HOME)
	rm -rf $(TARGET)
	cp -R $(SKILL_DIR) $(TARGET)
	rm -f $(TARGET)/Makefile
	@echo "Installed (copy): $(TARGET)"
	@echo "Restart your Claude Code session to pick it up."

uninstall:
	rm -rf $(TARGET)
	@echo "Removed $(TARGET)"

status:
	@ls -la $(TARGET) 2>/dev/null || echo "$(SKILL_NAME): not installed"
