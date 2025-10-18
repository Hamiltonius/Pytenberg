.PHONY: setup reset run

setup:
	python3 scripts/testbed.py --apply testbed.yaml

reset:
	python3 scripts/testbed.py --apply testbed.yaml --reset
	mkdir -p drop out logs refs

run:
	python3 pytenberg.py

