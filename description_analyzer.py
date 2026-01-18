# --- FILE: description_analyzer.py ---
import os
import re
import spacy
import pandas as pd
import numpy as np
import unicodedata
from PySide6.QtCore import QObject, Signal
from utils import resource_path  # A função resource_path é a chave
from collections import defaultdict, Counter

# --- Thresholds for Activity Analysis ---
MIN_VOTE_THRESHOLD = 1
CATEGORY_SIMILARITY_THRESHOLD = 0.60 

class DescriptionAnalyzer(QObject):
    progress = Signal(str)

    def __init__(self):
        super().__init__()
        self.nlp = None
        self.keyword_map = None 
        self.official_desc_docs = None
        
        self.non_informative_words = {
            'de', 'a', 'o', 'que', 'e', 'do', 'da', 'em', 'um', 'para', 'com', 'não', 'uma', 'os', 'no', 'na',
            'por', 'mais', 'as', 'dos', 'como', 'mas', 'ao', 'ele', 'das', 'à', 'seu', 'sua', 'ou', 'quando',
            'muito', 'nos', 'já', 'eu', 'também', 'só', 'pelo', 'pela', 'até', 'isso', 'ela', 'entre', 'depois',
            'sem', 'mesmo', 'aos', 'se', 'ter', 'serviço', 'serviços', 'nota', 'fiscal', 'valor', 'referente',
            'prestado', 'prestados', 'conforme', 'contrato', 'recebido', 'rec', 'fat', 'nf', 'nfs', 'nfe',
            'pagamento', 'pgto', 'mes', 'n', 'r', 'total', 'data', 'numero', 'discriminacao', 'desc', 'desc.'
        }
        
        self.curitiba_neighborhoods = {
            'abranches', 'agua verde', 'ahu', 'alto da gloria', 'alto da rua xv', 'augusta',
            'bacacheri', 'bairro alto', 'batel', 'boa vista', 'bom retiro', 'boqueirao',
            'butiatuvinha', 'cabral', 'cachoeria', 'cajuru', 'campina do siqueira',
            'campo comprido', 'campo de santana', 'capao da imbuia', 'capao raso',
            'cascatinha', 'centro', 'centro civico', 'cidade industrial de curitiba', 'cic',
            'cristo rei', 'fanny', 'fazendinha', 'ganchinho', 'hauer', 'hugo lange',
            'jardim botanico', 'jardim das americas', 'jardim social', 'juveve', 'lamenha pequena',
            'lindoia', 'mercer', 'merces', 'mossungue', 'novo mundo', 'orleans', 'parolin',
            'passauna', 'pilarzinho', 'pinheirinho', 'portao', 'prado velho', 'reboucas',
            'santa candida', 'santa felicidade', 'santa quiteria', 'santo inacio',
            'sao braz', 'sao francisco', 'sao joao', 'sao lourenco', 'sao miguel',
            'seminario', 'sitio cercado', 'taboao', 'taruma', 'tatuquara', 'tingui',
            'uberaba', 'umbara', 'vila izabel', 'vista alegre', 'xaxim'
        }
        self.known_brazilian_cities = {
            'rio branco', 'maceio', 'macapa', 'manaus', 'salvador', 'fortaleza', 'brasilia',
            'vitoria', 'goiania', 'sao luis', 'cuiaba', 'campo grande', 'belo horizonte',
            'belem', 'joao pessoa', 'recife', 'teresina', 'rio de janeiro',
            'natal', 'porto alegre', 'porto velho', 'boa vista', 'florianopolis', 'sao paulo',
            'aracaju', 'palmas', 'campinas', 'guarulhos', 'sao bernardo do campo', 'santo andre', 'osasco',
            'ribeirao preto', 'sao jose dos campos', 'sorocaba', 'duque de caxias',
            'nova iguacu', 'sao goncalo', 'niteroi', 'contagem', 'uberlandia', 'juiz de fora',
            'joinville', 'londrina', 'maringa', 'caxias do sul', 'pelotas', 'feira de santana',
            'campina grande', 'olinda', 'jaboatao dos guararapes', 'anapolis',
            'aparecida de goiania', 'vila velha', 'serra', 'campos dos goytacazes',
            'araucaria', 'sao jose dos pinhais', 'colombo', 'fazenda rio grande',
             'quatro barras', 'mandirituba', 'paranagua', 'guaratuba', 'matinhos',
              'morretes', 'contenda', 'balsa nova', 'campo largo', 'campo magro',
               'pinhais', 'piraquara', 'itaperucu', 'rio branco do sul', 'tijucas',
                'lapa', 'campo do tenente', 'pien', 'tunas', 'bocaiuva', 'sjp', 'cerro azul'
        }

    def load_models(self):
        try:
            if not self.nlp:
                self.progress.emit("🧠 Carregando modelo de linguagem (spaCy)...")
                
                # ✅ --- START: THIS IS THE CORRECT LOGIC ---
                # 1. Encontra a pasta md' na raiz do projeto
                #    (Funciona para 'python gui.py' e para o .exe)
                model_path = resource_path('pt_core_news_md')
                
                # 2. Carrega o modelo usando o caminho explícito
                self.nlp = spacy.load(model_path)
                # ✅ --- END: THIS IS THE CORRECT LOGIC ---
            
            self.progress.emit("✅ Modelo de IA (spaCy) carregado.")
            return True
        except OSError as e: 
            # ✅ --- MENSAGEM DE ERRO NOVA E CORRETA ---
            model_path_for_error = "ERRO: resource_path('pt_core_news_md') falhou"
            try:
                # Tenta construir o caminho para a mensagem de erro
                model_path_for_error = resource_path('pt_core_news_md')
            except:
                pass 
                
            self.progress.emit(f"❌ ERRO: Não foi possível carregar o spaCy.")
            self.progress.emit(f"   Verifique se a pasta 'pt_core_news_md' (com meta.json, vocab, etc. dentro) está na raiz do seu projeto.")
            self.progress.emit(f"   (Caminho tentado: {model_path_for_error})")
            return False
        except Exception as e:
            self.progress.emit(f"❌ ERRO ao carregar modelo spaCy: {e}")
            return False

    def _normalize_text(self, text):
        if not isinstance(text, str):
            return ""
        nfkd_form = unicodedata.normalize('NFD', text)
        return u"".join([c for c in nfkd_form if not unicodedata.combining(c)]).lower()

    def _find_service_locations(self, description, home_city="curitiba"):
        """
        Usa uma busca "brute-force" de keywords contra as listas de cidades.
        """
        if not isinstance(description, str):
            return []
            
        found_cities = set()
        
        normalized_description = self._normalize_text(description)
        normalized_home_city = self._normalize_text(home_city)
        normalized_neighborhoods = {self._normalize_text(n) for n in self.curitiba_neighborhoods}
        away_cities = self.known_brazilian_cities 

        for city in away_cities:
            try:
                if re.search(r'\b' + re.escape(city) + r'\b', normalized_description):
                    found_cities.add(city.title()) 
            except re.error:
                if city in normalized_description:
                    found_cities.add(city.title())

        final_cities_to_alert = set()
        for city_name in found_cities:
            normalized_city = self._normalize_text(city_name)
            if (normalized_city != normalized_home_city and
                normalized_city not in normalized_neighborhoods):
                final_cities_to_alert.add(city_name)
                    
        return list(final_cities_to_alert)

    def _get_key_lemmas(self, text):
        """
        Extrai os lemas-chave (substantivos, verbos, etc.)
        de um texto, removendo "ruído" genérico.
        """
        if not isinstance(text, str): 
            return []
        
        text_clean = re.sub(r'[^a-zA-Zà-úÇç\s]', ' ', text.lower())
        doc = self.nlp(text_clean)
        
        key_lemmas = []
        for token in doc:
            if token.pos_ in ['NOUN', 'PROPN', 'ADJ', 'VERB']:
                if (token.lemma_ not in self.non_informative_words and
                    token.has_vector and not token.is_oov):
                    key_lemmas.append(token.lemma_)
                    
        return list(set(key_lemmas))

    def _build_keyword_map(self, activity_data):
        """
        Cria o "Keyword Map" reverso E o cache de "Official Docs".
        """
        self.progress.emit("🗺️  Construindo mapa de palavras-chave e descrições...")
        keyword_map = defaultdict(list)
        official_desc_docs = {} 
        
        for code, entries in activity_data.items():
            for (description, aliquot, synonyms) in entries: 
                if code not in official_desc_docs:
                    official_desc_docs[code] = self.nlp(description)
                
                desc_lemmas = self._get_key_lemmas(description)
                for lemma in desc_lemmas:
                    if lemma not in keyword_map: 
                        keyword_map[lemma].append(code)
                
                if synonyms and isinstance(synonyms, str):
                    synonyms_clean = synonyms.replace(",", " ")
                    synonym_lemmas = self._get_key_lemmas(synonyms_clean)
                    for lemma in synonym_lemmas:
                        if code not in keyword_map[lemma]: 
                            keyword_map[lemma].append(code)
                        
        self.keyword_map = keyword_map
        self.official_desc_docs = official_desc_docs
        self.progress.emit("✅ Mapa de palavras-chave construído.")

    def analyze_invoices(self, df_invoices, activity_data):
        """
        Executa a análise de "3 Etapas" (Triage, Voting, Similarity)
        """
        if not self.load_models():
            df_invoices['location_alert'] = "Erro no Modelo NER"
            df_invoices['activity_alert'] = "Erro no Modelo NER"
            return df_invoices

        # --- 1. Location Analysis ---
        self.progress.emit("📍 Verificando localização do serviço...")
        locations = df_invoices['DISCRIMINAÇÃO DOS SERVIÇOS'].apply(lambda x: self._find_service_locations(x))
        df_invoices['location_alert'] = locations.apply(lambda x: ', '.join(x))
        
        # --- 2. "3-Stage" Activity Analysis ---
        
        if not self.keyword_map:
            self._build_keyword_map(activity_data)
            
        activity_alerts = []
        
        for index, row in df_invoices.iterrows():
            invoice_desc_str = row.get('DISCRIMINAÇÃO DOS SERVIÇOS')
            declared_code = row.get('CÓDIGO DA ATIVIDADE')
            
            # --- STAGE 1: TRIAGE ---
            key_lemmas = self._get_key_lemmas(invoice_desc_str)
            
            if not key_lemmas:
                activity_alerts.append("")
                continue
            
            # --- STAGE 2: KEYWORD VOTING ---
            code_votes = Counter()
            matched_keywords = []
            
            for lemma in key_lemmas:
                if lemma in self.keyword_map:
                    for code in self.keyword_map[lemma]:
                        code_votes[code] += 1
                    matched_keywords.append(lemma)
            
            if not code_votes:
                activity_alerts.append("") 
                continue

            winning_code, vote_count = code_votes.most_common(1)[0]
            
            if vote_count < MIN_VOTE_THRESHOLD:
                activity_alerts.append("") 
                continue
                
            if winning_code == declared_code:
                activity_alerts.append("") 
                continue

            # --- STAGE 3: SIMILARITY CHECK ---
            doc_declared = self.official_desc_docs.get(declared_code)
            doc_winner = self.official_desc_docs.get(winning_code)

            if not doc_declared or not doc_winner or not doc_declared.has_vector or not doc_winner.has_vector:
                activity_alerts.append("") 
                continue
                
            similarity_score = doc_declared.similarity(doc_winner)
            
            if similarity_score < CATEGORY_SIMILARITY_THRESHOLD:
                alert_text = (
                    f"Alerta: Descrição sugere Cód. '{winning_code}' "
                    f"(via: {', '.join(matched_keywords)}), "
                    f"que é semanticamente diferente do declarado ('{declared_code}')."
                )
                activity_alerts.append(alert_text)
            else:
                activity_alerts.append("")
                
        df_invoices['activity_alert'] = activity_alerts
        
        self.progress.emit("✅ Análise de IA concluída.")
        return df_invoices
