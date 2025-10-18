#!/bin/bash

VERSION=$(grep -oP '\$script:version\s*=\s*"\K[\d\.]+' MediaInspector.ps1)
ARCHIVE_NAME="MediaInspector-${VERSION}"

mkdir -p package

EXCLUDE_PATTERNS="-x!*.git -x!Profile -x!.gitignore -x!package.sh -x!package"

7z a -t7z -r -mx=9 "package/${ARCHIVE_NAME}.7z" ./* ${EXCLUDE_PATTERNS}

echo "---"
echo "package/${ARCHIVE_NAME}.7z"