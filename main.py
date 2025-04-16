from src.service.translation_service import translate_excel

if __name__ == '__main__':
    translation_dto_file_path = "resources/translation_dto_file.json"
    translate_excel(translation_dto_file_path)
