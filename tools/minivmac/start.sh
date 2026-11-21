#!/bin/bash
# Restart Mini vMac and mount a system disk, then any further disk images, in order.
#
#   tools/minivmac/start.sh SYSTEM.dsk [DISK.dsk ...]
#
# Set MINIVMAC to the app's path if it is not /Applications/Mini vMac.app. Mini vMac mounts at most six
# images. The system disk must finish booting before the others are mounted, hence the long first wait.
set -euo pipefail
app="${MINIVMAC:-/Applications/Mini vMac.app}"
[ $# -ge 1 ] || { echo "usage: $0 SYSTEM.dsk [DISK.dsk ...]" >&2; exit 2; }

pkill -f "Mini vMac.app/Contents/MacOS/minivmac" 2>/dev/null && sleep 2 || true
open -a "$app"
sleep 3
open -a "$app" "$1"
sleep 25
shift
for disk in "$@"; do
    open -a "$app" "$disk"
    sleep 4
done
