import docx2txt
from googletrans import Translator

def translate_word_file(input_file, output_file, target_language):
    # Read the content of the Word file
    try:
        text = docx2txt.process(input_file)
    except KeyError as e:
        print(f"Error reading the Word file: {e}. Ensure the file is a valid Word document.")
        return
    except Exception as e:
        print(f"Unexpected error reading the Word file: {e}")
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
        with open(output_file, 'w', encoding='utf-8') as output:
            output.write(translated_text)
        print(f"Translation saved to {output_file}")
    except Exception as e:
        print(f"Error saving the translated content: {e}")

if __name__ == "__main__":
    # Get input parameters from the command line
    input_file = input("Enter the path to the Word file: ").strip()
    output_file = input("Enter the path to save the translated Word file: ").strip()
    target_language = input("Enter the target language (e.g., 'fr' for French): ").strip()

    # Call the translation function
    translate_word_file(input_file, output_file, target_language)