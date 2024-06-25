def parse_json(json_str):
    stack = []
    result = {}
    current_key = None
    current_value = None
    in_string = False
    escape = False

    i = 0
    while i < len(json_str):
        char = json_str[i]

        if in_string:
            if escape:
                escape = False
            elif char == '\\':
                escape = True
            elif char == '"':
                in_string = False
            current_value += char
        else:
            if char == '{':
                if current_key is not None:
                    stack.append(result)
                    result[current_key] = {}
                    result = result[current_key]
                    current_key = None
                else:
                    stack.append({})
                    result = stack[-1]
            elif char == '}':
                if current_key is not None:
                    result[current_key] = current_value
                    current_key = None
                    current_value = None
                if stack:
                    result = stack.pop()
            elif char == '[':
                if current_key is not None:
                    stack.append(result)
                    result[current_key] = []
                    result = result[current_key]
                    current_key = None
                else:
                    stack.append([])
                    result = stack[-1]
            elif char == ']':
                if current_value is not None:
                    result.append(current_value)
                    current_value = None
                if stack:
                    result = stack.pop()
            elif char == '"':
                in_string = True
                current_value = ""
            elif char == ',':
                if current_value is not None:
                    if current_key is None:
                        result.append(current_value)
                    else:
                        result[current_key] = current_value
                    current_value = None
            elif char == ':':
                current_key = current_value
                current_value = None
            elif char == ' ' or char == '\n' or char == '\t' or char == '\r':
                pass  # ignore whitespace
            else:
                if current_value is None:
                    current_value = char
                else:
                    current_value += char

        i += 1

    if
