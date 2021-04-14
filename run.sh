#!/bin/bash

if [ ! -d venv ]; then
  python3 -m virtualenv venv
  source venv/bin/activate
  pip install -r requirements.txt
else
  source venv/bin/activate
fi

while read pub; do
  python3 pre_calc.py $pub;
done < publishers.txt