#!/bin/bash
D=/Users/muse/compare-vs-word/dialog
{ echo "pdf_cleanup at $(date)"; osascript "$D/pdf_cleanup.applescript" 2>&1; echo "rc=$?"; } > "$D/pdf_cleanup.out" 2>&1
