import math

def count_rendered_lines(text: str, chars_per_line: int = 0) -> int:
    """
    Count the number of rendered visual lines for a given text,
    taking into account hard returns (\n), soft returns (\x0b), 
    and line wrapping based on chars_per_line.
    
    If chars_per_line is 0 or less, wrapping is not calculated 
    (only explicit returns are counted).
    """
    if not text:
        return 1
        
    paragraphs = text.replace('\x0b', '\n').split('\n')
    total_lines = 0
    for p in paragraphs:
        if not p:
            total_lines += 1
            continue
            
        if chars_per_line <= 0:
            total_lines += 1
        else:
            total_lines += math.ceil(len(p) / chars_per_line)
            
    return total_lines
