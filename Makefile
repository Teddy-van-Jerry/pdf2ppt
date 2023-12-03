# The script to install 
SCRIPT=pdf2ppt

# The install directory
INSTALL_DIR=/usr/local/bin

install:
	@echo "Installing $(SCRIPT) to $(INSTALL_DIR)" 
	cp $(SCRIPT) $(INSTALL_DIR)
	chmod +x $(INSTALL_DIR)/$(SCRIPT)

uninstall:
	@echo "Removing $(SCRIPT) from $(INSTALL_DIR)"
	rm -f $(INSTALL_DIR)/$(SCRIPT)

.PHONY: install uninstall
