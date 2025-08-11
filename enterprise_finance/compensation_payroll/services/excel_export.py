# services/export_service.py

class ExportUtilityService:
    def split_header_to_lines(self, header_text, max_lines=3):
        clean_text = header_text.replace('_', ' ').title()
        words = clean_text.split()
        if len(words) <= max_lines:
            return '\n'.join(words)
        avg = len(words) // max_lines
        lines = []
        start = 0
        for i in range(max_lines):
            if i == max_lines - 1:
                line_words = words[start:]
            else:
                line_words = words[start:start+avg]
            lines.append(' '.join(line_words))
            start += avg
        return '\n'.join(lines)



