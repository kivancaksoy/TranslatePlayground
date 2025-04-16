class TranslationDto:
    def __init__(
            self,
            input_file,
            output_file,
            translate_service,
            deepl_api_key,
            source_lang,
            target_lang,
            columns_to_translate,
            all_columns
    ):
        self.input_file = input_file
        self.output_file = output_file
        self.translate_service = translate_service
        self.deepl_api_key = deepl_api_key
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.columns_to_translate = columns_to_translate
        self.all_columns = all_columns
