#!/bin/bash
for filename in owasp-mstg/Document/0x05*.md owasp-mstg/Document/0x06*.md; do
    docker run --rm -u `id -u`:`id -g` -v `pwd`:/pandoc dalibo/pandocker --section-divs -f markdown -t html $filename -o $(basename $filename .md).html
done

mkdir -p tools/html
mv *.html tools/html