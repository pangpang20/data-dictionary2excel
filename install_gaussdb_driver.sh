#!/bin/bash
# install_gaussdb_driver.sh
# Automatically download, install, and configure GaussDB driver, supporting HCE, CentOS (Hce2), Euler, Kylin systems
# Idempotent and repeatable execution

set -euo pipefail

#===================
#     Basic Configuration     
#===================
DOWNLOAD_URL="https://dbs-download.obs.cn-north-1.myhuaweicloud.com/GaussDB/1730887196055/GaussDB_driver.zip"
HOME_DIR="$HOME"
ZIP_FILE="$HOME_DIR/GaussDB_driver.zip"
DRIVER_DIR="$HOME_DIR/GaussDB_driver/Centralized"
LIB_DIR="$HOME_DIR/GaussDB_driver_lib"
LOG_FILE="/tmp/gaussdb_driver_install_$(date +%Y%m%d_%H%M%S).log"

#===================
#     Utility Functions     
#===================
log() { echo "[$(date '+%Y-%m-%d %H:%M:%S')] $*" | tee -a "$LOG_FILE"; }

cleanup() {
    log "Cleaning up temporary files..."
    [[ -f "$ZIP_FILE" ]] && rm -rf "$ZIP_FILE" 2>/dev/null
    [[ -d "$HOME_DIR/GaussDB_driver" ]] && rm -rf "$HOME_DIR/GaussDB_driver" 2>/dev/null
}

#===================
#     Parameter Checks     
#===================
command -v wget >/dev/null || { log "Error: wget is missing"; exit 1; }
command -v unzip >/dev/null || { log "Error: unzip is missing"; exit 1; }
command -v tar >/dev/null || { log "Error: tar is missing"; exit 1; }
command -v ldconfig >/dev/null || { log "Error: ldconfig is missing"; exit 1; }

log "Starting GaussDB driver installation..."

#===================
#     Download and Extract   
#===================
if [[ ! -f "$ZIP_FILE" ]]; then
    log "Downloading driver file..."
    wget -O "$ZIP_FILE" "$DOWNLOAD_URL" >> "$LOG_FILE" 2>&1 || { log "Error: Download failed"; exit 1; }
else
    log "Driver file already exists, skipping download"
fi

log "Extracting driver file..."
unzip -o "$ZIP_FILE" -d "$HOME_DIR/" >> "$LOG_FILE" 2>&1 || { log "Error: Extraction failed"; exit 1; }


#===================
#     Detect System and Architecture   
#===================
ARCH=$(uname -m)
case "$ARCH" in
    x86_64) ARCH_TYPE="X86_64" ;;
    aarch64) ARCH_TYPE="arm_64" ;;
    *) log "Error: Unsupported architecture: $ARCH"; exit 1 ;;
esac
log "architecture: $ARCH_TYPE"
OS_TYPE=""
if [[ -f /etc/os-release ]]; then
    . /etc/os-release
    case "$ID" in
        centos|hce)
            if [[ -d "$DRIVER_DIR/Hce2_$ARCH_TYPE" ]]; then
                OS_TYPE="Hce2_$ARCH_TYPE"
            fi
            ;;
        euler)
            VERSION=$(grep -oP 'VERSION_ID="\K[^"]+' /etc/os-release)
            case "$VERSION" in
                2.5*)
                    if [[ -d "$DRIVER_DIR/Euler2.5_$ARCH_TYPE" ]]; then
                        OS_TYPE="Euler2.5_$ARCH_TYPE"
                    fi
                    ;;
                2.9*)
                    if [[ -d "$DRIVER_DIR/Euler2.9_$ARCH_TYPE" ]]; then
                        OS_TYPE="Euler2.9_$ARCH_TYPE"
                    fi
                    ;;
            esac
            ;;
        kylin)
            if [[ -d "$DRIVER_DIR/Kylinv10_$ARCH_TYPE" ]]; then
                OS_TYPE="Kylinv10_$ARCH_TYPE"
            fi
            ;;
        *)
            log "Error: Unsupported operating system: $ID"; exit 1
            ;;
    esac
else
    log "Warning: Unable to read /etc/os-release, attempting to infer system type from directory structure"
    if [[ -d "$DRIVER_DIR/Hce2_$ARCH_TYPE" ]]; then
        OS_TYPE="Hce2_$ARCH_TYPE"
    elif [[ -d "$DRIVER_DIR/Euler2.5_$ARCH_TYPE" ]]; then
        OS_TYPE="Euler2.5_$ARCH_TYPE"
    elif [[ -d "$DRIVER_DIR/Euler2.9_$ARCH_TYPE" ]]; then
        OS_TYPE="Euler2.9_$ARCH_TYPE"
    elif [[ -d "$DRIVER_DIR/Kylinv10_$ARCH_TYPE" ]]; then
        OS_TYPE="Kylinv10_$ARCH_TYPE"
    else
        log "Error: Unsupported operating system or architecture: $ARCH_TYPE"; exit 1
    fi
fi
log "Detected system: $OS_TYPE"
if [[ -z "$OS_TYPE" ]]; then
    log "Error: No matching driver directory found: $DRIVER_DIR/*_$ARCH_TYPE"; exit 1
fi



#===================
#     Copy Driver Package   
#===================
mkdir -p "$LIB_DIR"
DRIVER_PACKAGE=$(find "$DRIVER_DIR/$OS_TYPE" -name "*Python.tar.gz" | head -n 1)
if [[ -z "$DRIVER_PACKAGE" ]]; then
    log "Error: No driver package found for $OS_TYPE"; exit 1
fi

log "Copying driver package: $DRIVER_PACKAGE to $LIB_DIR"
sudo cp "$DRIVER_PACKAGE" "$LIB_DIR/" || { log "Error: Failed to copy driver package"; exit 1; }

#===================
#     Extract Driver Package   
#===================
log "Extracting driver package to $LIB_DIR..."
tar -zxvf "$LIB_DIR/$(basename "$DRIVER_PACKAGE")" -C "$LIB_DIR/" >> "$LOG_FILE" 2>&1 || { log "Error: Failed to extract driver package"; exit 1; }
rm -f "$LIB_DIR/$(basename "$DRIVER_PACKAGE")"
sudo chmod 755 -R $LIB_DIR

#===================
#     Configure Dynamic Link Library   
#===================
log "Configuring dynamic link library path..."
echo "$LIB_DIR/lib" | sudo tee /etc/ld.so.conf.d/gauss-libpq.conf >/dev/null
if ! grep -Fx "$LIB_DIR/lib" /etc/ld.so.conf >/dev/null; then
    sudo sed -i "1s|^|$LIB_DIR/lib\n|" /etc/ld.so.conf
fi
sudo sed -i '/gauss/d' /etc/ld.so.conf
sudo ldconfig



#===================
#     Verify Installation     
#===================
if sudo ldconfig -p | grep -q libpq; then
	cleanup
    log "============================================================="
    log "GaussDB driver installed successfully!"
    log "Dynamic link library configured: $LIB_DIR/lib"
    log "Log file: $LOG_FILE"
    log "============================================================="
else
    log "Error: Dynamic link library verification failed"
    exit 1
fi
