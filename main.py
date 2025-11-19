import csv
import uuid
import sys
import argparse
from datetime import datetime, timedelta
from typing import List, Optional
import logging

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def parse_date(date_str: str) -> Optional[datetime]:
    """Tenta converter string para datetime (DD/MM/YYYY ou YYYY-MM-DD)."""
    for fmt in ('%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None

def format_ics_date(dt: datetime) -> str:
    return dt.strftime('%Y%m%dT%H%M%S')

def generate_event_block(subject: str, start_date: datetime) -> List[str]:
    """Gera o bloco do evento genérico."""
    
    # Define horário padrão: 09:00 às 09:30
    start_event = start_date.replace(hour=9, minute=0, second=0)
    end_event = start_event + timedelta(minutes=30)
    
    return [
        "BEGIN:VEVENT",
        f"UID:{uuid.uuid4()}",
        f"DTSTAMP:{format_ics_date(datetime.now())}",
        f"DTSTART;TZID=America/Sao_Paulo:{format_ics_date(start_event)}",
        f"DTEND;TZID=America/Sao_Paulo:{format_ics_date(end_event)}",
        f"SUMMARY:{subject}",
        "STATUS:CONFIRMED",
        "CLASS:PUBLIC",
        "BEGIN:VALARM",
        "ACTION:DISPLAY",
        "TRIGGER;RELATED=START:-PT15M",
        "DESCRIPTION:Lembrete gerado via script CSV",
        "END:VALARM",
        "END:VEVENT"
    ]

def generate_ics_structure(events_content: List[str]) -> List[str]:
    header = [
        "BEGIN:VCALENDAR",
        "PRODID:-//CSV-to-ICS//Github//PT-BR",
        "VERSION:2.0",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:Eventos Importados",
        "BEGIN:VTIMEZONE",
        "TZID:America/Sao_Paulo",
        "BEGIN:STANDARD",
        "DTSTART:16010101T000000",
        "TZOFFSETTO:-0300",
        "TZOFFSETFROM:-0300",
        "TZNAME:-03",
        "END:STANDARD",
        "END:VTIMEZONE"
    ]
    return header + events_content + ["END:VCALENDAR"]

def process_csv(input_file: str, output_file: str, delimiter: str, 
                col_subject: str, col_date: str, default_subject: str):
    
    ics_events = []
    processed_count = 0
    
    try:
        # Tenta abrir com utf-8, fallback para cp1252
        try:
            file_obj = open(input_file, newline='', encoding='utf-8')
            reader = csv.DictReader(file_obj, delimiter=delimiter)
            next(reader); file_obj.seek(0)
            reader = csv.DictReader(file_obj, delimiter=delimiter)
        except UnicodeDecodeError:
            file_obj = open(input_file, newline='', encoding='cp1252')
            reader = csv.DictReader(file_obj, delimiter=delimiter)

        with file_obj:
            # Normaliza nomes das colunas do CSV para minúsculo
            fieldnames = [f.strip().lower() for f in reader.fieldnames] if reader.fieldnames else []
            target_col_subj = col_subject.lower()
            target_col_date = col_date.lower()

            for row in reader:
                row_clean = {k.strip().lower(): v.strip() for k, v in row.items() if k}

                # ==============================================================================
                # LÓGICA DE VALOR PADRÃO / COLUNA INEXISTENTE
                # Se a coluna de assunto não for encontrada ou estiver vazia, usa o valor padrão
                # ==============================================================================
                subject_value = row_clean.get(target_col_subj)
                
                if not subject_value:
                    # Aqui aplicamos o valor padrão passado por argumento
                    subject_value = default_subject
                    logging.warning(f"Assunto vazio na linha. Usando padrão: '{default_subject}'")
                # ==============================================================================

                date_str = row_clean.get(target_col_date)

                if not date_str:
                    logging.warning(f"Data não encontrada. Pulando linha: {row}")
                    continue

                dt_event = parse_date(date_str)
                if not dt_event:
                    logging.error(f"Data inválida '{date_str}'. Use DD/MM/AAAA ou AAAA-MM-DD.")
                    continue

                ics_events.extend(generate_event_block(subject_value, dt_event))
                processed_count += 1

        full_content = generate_ics_structure(ics_events)

        with open(output_file, 'w', encoding='utf-8', newline='') as f:
            f.write('\r\n'.join(full_content))

        logging.info(f"Sucesso! {processed_count} eventos criados em '{output_file}'.")

    except FileNotFoundError:
        logging.critical(f"Arquivo '{input_file}' não encontrado.")
    except Exception as e:
        logging.critical(f"Erro: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Converte CSV para Calendário ICS.")
    parser.add_argument("-i", "--input", required=True, help="Arquivo CSV de entrada.")
    parser.add_argument("-o", "--output", default="calendario.ics", help="Arquivo ICS de saída.")
    parser.add_argument("-d", "--delimiter", default=";", help="Delimitador do CSV.")
    
    # Argumentos para personalizar colunas
    parser.add_argument("--col-subject", default="Assunto", help="Nome da coluna do título do evento.")
    parser.add_argument("--col-date", default="Data", help="Nome da coluna da data do evento.")
    
    # ==============================================================================
    # CONFIGURAÇÃO DO VALOR PADRÃO NO ARGUMENTO
    # Permite definir o nome padrão caso a coluna não exista no CSV
    # ==============================================================================
    parser.add_argument("--default-subject", default="Evento Importado", 
                        help="Nome padrão para o evento caso a coluna esteja vazia.")
    
    args = parser.parse_args()

    process_csv(args.input, args.output, args.delimiter, 
                args.col_subject, args.col_date, args.default_subject)