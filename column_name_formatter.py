""" Map of raw names to standard names """
_name_cache = {}


def fmt(column_name):
    """
    Format a column name so it's consistent across different sources

    Column names may not match exactly between excel sheets and word docs. This formats the input strings so that they
    will match. Most alterations are generic; removing all white space and new lines, lower case, percentages changed to
    %, but some are specific, such as 'over 60' needing to be mapped to 60+ for example

    Args:
        column_name: name of the column to format

    Returns:
        standardised formatted column name
    """
    if column_name in _name_cache:
        out = _name_cache[column_name]
    else:
        out = str.lower(column_name)            # lower case
        out = out.replace("\n", "")             # remove newlines
        out = "".join(out.split(" "))           # remove all whitespace
        out = out.replace("*", "")              # to remove any notes
        out = out.replace("percentage", "%")    # handle %age signs
        out = out.replace("no.", "number")      # default all no. to number
        out = out.replace("\u2013", "-")        # spaces around dashes in MS word convert them to a endash
        out = out.replace("under25", "u25")    # special case
        out = out.replace("60plus", "60+")     # special case
        out = out.replace("60&over", "60+")   # special case

        _name_cache[column_name] = out

    return out
