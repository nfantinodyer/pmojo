import re
with open('practicemojo.txt', 'r', encoding='utf-8') as text_file:
    text = text_file.readlines()

# Remove header and footer lines if not part of the needed data
del text[0:8]
del text[-5:]

print(text)

appointments = {}

for line in text:
    if re.search("Family",line):  # Skip the line containing "Family" as it's not needed
        continue
    # Normalize spaces and remove non-ASCII characters
    line = re.sub(r"\s+", ' ', line.strip())
    line = re.sub(r"[^\x00-\x7F]+", '', line)

    # Match the expected line format with dates and times
    match = re.search(r"([a-zA-Z, ]+)(\d{2}/\d{2} @ \d{2}:\d{2} [APM]+)", line)
    if match:
        name, datetime = match.groups()
        date, time = datetime.split(' @ ')
        # Construct a unique key for each person and date
        key = f"{name.strip()}{date}"
        if key not in appointments:
            appointments[key] = []
        appointments[key].append(time)
    else:
        print(f"No match for line: {line}")  # Debug output for unmatched lines

# Prepare to write consolidated entries to file
lines_to_write = []
for key, times in appointments.items():
    # Combine multiple times into one line
    times_str = ' & '.join(sorted(set(times)))  # Remove duplicates and sort times
    # Use regex to extract the date part
    match = re.search(r"(\d{2}/\d{2})$", key)
    if match:
        date = match.group(1)
        name = key[:match.start()].strip()
        line_to_write = f"{name} {date} @ {times_str}"
        lines_to_write.append(line_to_write)
    else:
        print(f"Unexpected key format: {key}")  # Debug output for unexpected key format


print(lines_to_write)