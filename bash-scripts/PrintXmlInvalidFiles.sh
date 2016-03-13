#!/bin/bash
# find invalid xml files recursivly, print filenames
if [[ $# -ne 1 ]] 
then
    echo "Usage: $0 Directory"
    exit 1
fi

find $1 -iname "*.xml" -print0 | xargs -0  -i xmlstarlet val {} | grep invalid | rev | cut -d - -f2- | rev | sed 's/ $//g'
