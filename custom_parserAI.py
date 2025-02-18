import re

def parse_command(command):
    command_lower = command.lower()
    
    # --- Sorting: Catch various phrasings for sorting ---
    sort_patterns = [
        # Pattern 1: "sort ... by Score in ascending/descending order"
        r"sort(?:.*?numerically)?\s+(?:this\s+file\s+)?by\s+(\w+).*?(ascending|asc|descending|desc)",
        # Pattern 2: "sort the file from low to high/by high to low"
        r"sort(?:.*?file)?\s+from\s+(low to high|high to low|highest to lowest|lowest to highest).*?by\s+(\w+)",
        # New Pattern 3: "Make it so that the data is sorted from high to low by score"
        r"make\s+it\s+so\s+that\s+the\s+data\s+is\s+sorted\s+from\s+(high to low|lowest to highest|highest to lowest)\s+by\s+(\w+)",
        # New Pattern 4: "I want you to sort this file by Score from high to low"
        r"i\s+want\s+you\s+to\s+sort\s+this\s+file\s+by\s+(\w+)\s+from\s+(high to low|highest to lowest|lowest to highest)"
    ]
    
    for pattern in sort_patterns:
        match = re.search(pattern, command_lower)
        if match:
            if len(match.groups()) == 2:
                # Depending on pattern structure, decide which group is header/order
                group1, group2 = match.group(1), match.group(2)
                # Determine if the first group is header or an order phrase
                if group1 in ["low to high", "high to low", "highest to lowest", "lowest to highest"]:
                    order = "asc" if group1 == "low to high" or group1 == "lowest to highest" else "desc"
                    header = group2
                else:
                    header = group1
                    order = "asc" if group2 in ["ascending", "asc", "low to high"] else "desc"
                return {
                    "function": "number_sorting",  # or "word_sorting" if you add further logic
                    "parameters": {"header": header.capitalize(), "order": order}
                }
    
    # --- Filtering: Handling various phrasings ---
    filter_patterns = [
        # Pattern: "filter the file where Score is greater than 50"
        r"filter.*?where\s+(\w+)\s+(?:is\s+)?(greater than|less than|above|below)\s+([\d\.]+)",
        # New Pattern 1: "Filter the file by score greater than 30"
        r"filter.*?by\s+(\w+)\s+(greater than|above)\s+([\d\.]+)",
        # New Pattern 2: "I want you to filter ..."
        r"i\s+want\s+you\s+to\s+filter.*?where\s+(\w+)\s+(?:is\s+)?(greater than|less than|above|below)\s+([\d\.]+)"
    ]
    
    for pattern in filter_patterns:
        match = re.search(pattern, command_lower)
        if match:
            header, condition, value = match.group(1), match.group(2), match.group(3)
            # Decide sort_order based on condition wording
            sort_order = "desc" if condition in ["greater than", "above"] else "asc"
            return {
                "function": "file_filter",
                "parameters": {"column": header.capitalize(), "value": value, "sort_order": sort_order}
            }
    
    # --- Duplicate Removal ---
    dup_pattern = r"(remove|delete).*(duplicate)"
    if re.search(dup_pattern, command_lower):
        return {"function": "dup_remover", "parameters": {}}
    
    # --- Number Highlighting ---
    highlight_patterns = [
        # Pattern: "highlight rows where Score is above 80"
        r"highlight.*?where\s+(\w+)\s+(?:is\s+)?(above|below|greater than|less than)\s+([\d\.]+)",
    ]
    
    for pattern in highlight_patterns:
        match = re.search(pattern, command_lower)
        if match:
            header, condition_word, value = match.group(1), match.group(2), match.group(3)
            condition = "desc" if condition_word in ["above", "greater than"] else "asc"
            return {
                "function": "num_highlight",
                "parameters": {"header": header.capitalize(), "condition": condition, "key_value": value}
            }
    
    # --- Word Highlighting ---
    word_highlight_pattern = r"highlight.*?where\s+(\w+)\s+(?:contains|has)\s+(\w+)"
    match = re.search(word_highlight_pattern, command_lower)
    if match:
        header, key_value = match.group(1), match.group(2)
        return {
            "function": "word_highlight",
            "parameters": {"header": header.capitalize(), "key_value": key_value}
        }

    word_sorting_patterns = [
        # Pattern 1: "Sort this file by [Header] from high to low"
        r"sort.*?this\s+file\s+by\s+(\w+)\s+(from\s+high\s+to\s+low|highest\s+to\s+lowest|low\s+to\s+high|ascending|descending)",
        # Pattern 2: "I want you to sort the file by [Header] from low to high"
        r"i\s+want\s+you\s+to\s+sort\s+this\s+file\s+by\s+(\w+)\s+(from\s+high\s+to\s+low|lowest\s+to\s+highest|ascending|descending)"
    ]

    for pattern in word_sorting_patterns:
        match = re.search(pattern, command_lower)
        if match:
            header, order_phrase = match.group(1), match.group(2)
            
            # Determine order (ascending or descending)
            if order_phrase in ["from high to low", "descending"]:
                order = "desc"
            elif order_phrase in ["from low to high", "ascending", "lowest to highest"]:
                order = "asc"
            else:
                order = "asc"  # Default to ascending if not specified
            
            return {
                "function": "word_sorting",
                "parameters": {"header": header.capitalize(), "order": order}
            }

    styling_patterns = [
        # Pattern 1: "Make it so that the headers are bold and have a blue background"
        r"make\s+it\s+so\s+that\s+the\s+headers\s+are\s+(bold)\s+and\s+have\s+a\s+(\w+)\s+background",
        # Pattern 2: "I want you to change the font color to red and add borders with thin/medium/thick thickness"
        r"i\s+want\s+you\s+to\s+change\s+the\s+font\s+color\s+to\s+(\w+)\s+and\s+add\s+borders\s+with\s+(thin|medium|thick)\s+thickness",
        # Pattern 3: "Make the headers bold with left alignment and the data right-aligned"
        r"make\s+the\s+headers\s+(bold)\s+with\s+(\w+)\s+alignment\s+and\s+the\s+data\s+(\w+)-aligned"
    ]

    for pattern in styling_patterns:
        match = re.search(pattern, command_lower)
        if match:
            if pattern == styling_patterns[0]:
                # Handle bold headers and background color
                bold_headers = True
                header_background = match.group(2)
                return {
                    "function": "styling",
                    "parameters": {
                        "bold_headers": bold_headers,
                        "header_background": header_background,
                    }
                }
            elif pattern == styling_patterns[1]:
                # Handle font color and border thickness (thin, medium, thick)
                font_color = match.group(1)
                border_thickness = match.group(2)
                return {
                    "function": "styling",
                    "parameters": {
                        "font_color": font_color,
                        "border_thickness": border_thickness,
                    }
                }
            elif pattern == styling_patterns[2]:
                # Handle bold headers, alignment of headers and data
                bold_headers = True
                header_alignment = match.group(2)
                data_alignment = match.group(3)
                return {
                    "function": "styling",
                    "parameters": {
                        "bold_headers": bold_headers,
                        "header_alignment": header_alignment,
                        "data_alignment": data_alignment,
                    }
                }

    return None

# Example usage:
if __name__ == "__main__":
    test_commands = [
        "Please sort this file numerically by score in ascending order",
        "Sort the file from high to low by Score",
        "Sort this file from highest to lowest by Score",
        "Make it so that the data is sorted from high to low by score",
        "I want you to sort this file by Score from high to low",
        
        "Filter the file where Age is greater than 30",
        "Filter the file by score greater than 30",
        "I want you to filter the file where Age is greater than 30",
        "Remove duplicate rows",

        "Highlight rows where Score is above 80",
        "Highlight rows where Name contains John"

        "Sort this file by score from high to low",
        "I want you to sort this file by name from low to high"

        "Make it so that the headers are bold and have a blue background",
        "I want you to change the font color to red and add borders with thin thickness",
        "Make the headers bold with left alignment and the data right-aligned"
    ]
    
    for cmd in test_commands:
        result = parse_command(cmd)
        print(f"Input: {cmd}")
        if result:
            print("Parsed as:", result, "\n")
        else:
            print("No matching command found.\n")
