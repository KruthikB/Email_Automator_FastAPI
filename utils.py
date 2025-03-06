def convert_to_str(value):
    if isinstance(value, int) or isinstance(value,float):
        return str(int(value))
    elif isinstance(value, str):
        if value.isdigit():
            return value
        try:
            float_value = str(value)
            if float(float_value).is_integer():
                return str(float_value[:float_value.index(".")]) if '.' in value else value
            else:
                return value
        except ValueError:
            return value
    else:
        return str(value)
    