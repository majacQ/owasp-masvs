#!/bin/bash
LanguageArray=("de" "en" "es" "fa" "fr" "hi" "ja" "ko" "ru" "zhcn" "zhtw")
for lang in ${LanguageArray[*]}; do
    # python3 export.py -f yaml -l $lang > masvs_$lang.yaml
    python3 parse_html.py -m masvs_$lang.yaml -i html -o masvs_full_$lang.yaml
    python3 yaml_to_excel.py -m masvs_full_$lang.yaml -i checklist.xlsx -o checklist_$lang.xlsx
done
