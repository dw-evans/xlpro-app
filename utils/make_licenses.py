import subprocess


# subprocess.run(
#     """pip-licenses \
#   --format=json \
#   --with-authors \
#   --with-urls \
#   --with-license-file \
#   --output-file=THIRD_PARTY_LICENSES.json \
#   --ignore-packages pip-licenses prettytable tomli wcwidth
# """,
#     capture_output=True,
#     check=True
# )
result = subprocess.run(
    "pip-licenses \
    --ignore-packages pip-licenses prettytable tomli wcwidth win32typelibs xlpro xlpro-cli \
    --format=markdown \
    --with-urls \
    --output-file=third_party_licenses_summary_raw.md",
    capture_output=True,
    # --format=plain-vertical \
    # --fail-on=MIT \
    # --partial-match \
)

print(result.stdout)
print(result.stderr)


# apache licenses only
    # packaging python-dateutil tzdata \
result = subprocess.run(
    "pip-licenses\
    --format=plain-vertical \
    --with-urls \
    --with-authors \
    --partial-match \
    --with-license-file \
    --output-file=third_party_licenses_raw.txt",
    capture_output=True,
    check=True
    # --format=plain-vertical \
    # --fail-on=MIT \
)

print(result.stdout)
print(result.stderr)


pass

