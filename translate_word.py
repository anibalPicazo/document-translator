import docx
from googletrans import Translator

def translate_word_file(input_file, output_file, target_language):
    # Read the content of the Word file
    try:
        doc = docx.Document(input_file)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        print(f"Error reading the Word file: {e}")
        return

    # Translate the content
    translator = Translator()
    try:
        translated_text = translator.translate(text, dest=target_language).text
    except Exception as e:
        print(f"Error translating the content: {e}")
        return

    # Save the translated text to a new Word file
    try:
        translated_doc = docx.Document()
        for translated_paragraph in translated_text.split("\n"):
            translated_doc.add_paragraph(translated_paragraph)
        translated_doc.save(output_file)
        print(f"Translation saved to {output_file}")
    except Exception as e:
        print(f"Error saving the translated content: {e}")

if __name__ == "__main__":
    # Get input parameters from the command line
    input_file = input("Enter the path to the Word file: ").strip()
    output_file = input("Enter the path to save the translated Word file (include .docx extension): ").strip()
    target_language = input("Enter the target language (e.g., 'fr' for French): ").strip()

    # Call the translation function
    translate_word_file(input_file, output_file, target_language)