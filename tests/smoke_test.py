import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app.server import ContractService, InvoiceService, ValidationService, ChatService


def main() -> None:
    contract_1 = ROOT / 'BWC Management Services Contract rates.xlsx'
    contract_2 = ROOT / 'Additional EDF contract rates.xlsx'
    invoice_pdf = ROOT / 'Bills' / '000027224635.pdf'

    print('Upserting contract 1...')
    print(ContractService.upsert_from_excel(contract_1))

    print('Upserting contract 2...')
    print(ContractService.upsert_from_excel(contract_2))

    print('Parsing invoice...')
    invoice = InvoiceService.parse_pdf(invoice_pdf)
    print('Invoice number:', invoice.get('invoice_number'))
    print('MPAN count:', len(invoice.get('mpans', {})))

    print('Running validation...')
    result = ValidationService.validate_invoice_record(invoice)
    print('Status:', result.get('status'))
    print('Score:', result.get('score'), result.get('score_band'))
    print('Reasons:', len(result.get('reasons', [])))

    print('Chat sample...')
    answer = ChatService.answer('What is the validation status and score?', invoice.get('invoice_number'))
    print(answer.get('answer'))
    print('Citations:', answer.get('citations'))


if __name__ == '__main__':
    main()
