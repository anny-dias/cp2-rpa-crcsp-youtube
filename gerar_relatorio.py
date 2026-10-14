"""Gera a camada gerencial sem rede e sem reprocessar as abas originais."""
import argparse
from pathlib import Path

from openpyxl import load_workbook
from relatorio_gerencial import build_report
from saida_segura import criar_temporario, publicar_excel


def gerar(entrada, saida=None):
    entrada = Path(entrada).resolve()
    saida = Path(saida).resolve() if saida else entrada.with_name(entrada.stem + '_GERENCIAL.xlsx')
    if entrada == saida:
        raise ValueError('Use uma saída diferente para preservar o Excel de origem.')
    saida.parent.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(entrada)
    temp = None
    try:
        build_report(wb, entrada.name)
        temp = criar_temporario(saida)
        wb.save(temp)
        saida = publicar_excel(temp, saida)
    finally:
        wb.close()
    return saida


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('arquivo', nargs='?')
    parser.add_argument('--saida')
    args = parser.parse_args()
    try:
        entrada = args.arquivo or input('Cole o caminho do Excel de coleta: ').strip().strip('\"')
        print(gerar(entrada, args.saida))
        return 0
    except (ValueError, OSError, EOFError) as exc:
        parser.error(str(exc))


if __name__ == '__main__':
    raise SystemExit(main())
