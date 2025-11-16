#!/usr/bin/env python3
"""
Script de validación de integridad para archivos financieros
Genera checksums SHA256 para trazabilidad audit-ready

Uso:
    python scripts/audit_integrity_check.py
    python scripts/audit_integrity_check.py --verify audit_checksums.json
"""

import hashlib
import json
from pathlib import Path
from datetime import datetime
import argparse
import sys


def calculate_checksum(file_path):
    """
    Calcula SHA256 de un archivo

    Args:
        file_path: Path del archivo a procesar

    Returns:
        str: Hash SHA256 en formato hexadecimal
    """
    sha256 = hashlib.sha256()
    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(4096), b''):
            sha256.update(chunk)
    return sha256.hexdigest()


def get_file_info(file_path):
    """
    Obtiene información completa del archivo

    Args:
        file_path: Path del archivo

    Returns:
        dict: Información del archivo (checksum, tamaño, fecha modificación)
    """
    stat = file_path.stat()
    return {
        'checksum': calculate_checksum(file_path),
        'size_bytes': stat.st_size,
        'modified': datetime.fromtimestamp(stat.st_mtime).isoformat(),
        'created': datetime.now().isoformat()
    }


def create_snapshot(data_dir='data', output_file='audit_checksums.json'):
    """
    Genera snapshot de checksums para auditoría

    Args:
        data_dir: Directorio con archivos a auditar
        output_file: Archivo JSON de salida
    """
    data_path = Path(data_dir)

    if not data_path.exists():
        print(f"❌ Error: Directorio '{data_dir}' no existe")
        print(f"💡 Tip: Crear estructura con: mkdir -p {data_dir}")
        sys.exit(1)

    snapshot = {
        'metadata': {
            'timestamp': datetime.now().isoformat(),
            'data_directory': str(data_path.absolute()),
            'version': '1.0'
        },
        'files': {}
    }

    # Procesar archivos Excel, CSV y otros financieros
    extensions = ['*.xlsx', '*.xls', '*.csv', '*.xlsm']
    total_files = 0

    print(f"📊 Generando snapshot de integridad...")
    print(f"📁 Directorio: {data_path.absolute()}")
    print(f"🔍 Extensiones: {', '.join(extensions)}\n")

    for pattern in extensions:
        for file_path in data_path.rglob(pattern):
            # Ignorar archivos temporales de Excel (~$)
            if file_path.name.startswith('~$'):
                continue

            rel_path = str(file_path.relative_to(data_path))
            print(f"  ✓ Procesando: {rel_path}")

            snapshot['files'][rel_path] = get_file_info(file_path)
            total_files += 1

    # Guardar snapshot
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(snapshot, f, indent=2, ensure_ascii=False)

    print(f"\n✅ Snapshot creado exitosamente")
    print(f"📄 Archivo: {output_file}")
    print(f"📊 Total archivos procesados: {total_files}")
    print(f"⏰ Timestamp: {snapshot['metadata']['timestamp']}")

    return snapshot


def verify_snapshot(snapshot_file='audit_checksums.json', data_dir='data'):
    """
    Verifica integridad comparando con snapshot anterior

    Args:
        snapshot_file: Archivo JSON con checksums de referencia
        data_dir: Directorio con archivos a verificar

    Returns:
        bool: True si todos los archivos son íntegros, False si hay discrepancias
    """
    snapshot_path = Path(snapshot_file)

    if not snapshot_path.exists():
        print(f"❌ Error: Snapshot '{snapshot_file}' no existe")
        print(f"💡 Tip: Crear snapshot con: python {__file__}")
        sys.exit(1)

    # Cargar snapshot de referencia
    with open(snapshot_path, 'r', encoding='utf-8') as f:
        reference = json.load(f)

    print(f"🔍 Verificando integridad contra snapshot...")
    print(f"📄 Snapshot: {snapshot_file}")
    print(f"⏰ Creado: {reference['metadata']['timestamp']}\n")

    data_path = Path(data_dir)
    all_ok = True
    checked = 0

    for rel_path, ref_info in reference['files'].items():
        file_path = data_path / rel_path

        if not file_path.exists():
            print(f"  ❌ FALTA: {rel_path}")
            all_ok = False
            continue

        current_checksum = calculate_checksum(file_path)

        if current_checksum == ref_info['checksum']:
            print(f"  ✅ OK: {rel_path}")
        else:
            print(f"  ⚠️  MODIFICADO: {rel_path}")
            print(f"      Esperado: {ref_info['checksum']}")
            print(f"      Actual:   {current_checksum}")
            all_ok = False

        checked += 1

    # Detectar archivos nuevos
    extensions = ['*.xlsx', '*.xls', '*.csv', '*.xlsm']
    for pattern in extensions:
        for file_path in data_path.rglob(pattern):
            if file_path.name.startswith('~$'):
                continue
            rel_path = str(file_path.relative_to(data_path))
            if rel_path not in reference['files']:
                print(f"  ℹ️  NUEVO: {rel_path}")

    print(f"\n{'='*60}")
    if all_ok:
        print(f"✅ VERIFICACIÓN EXITOSA")
        print(f"📊 {checked} archivos verificados sin cambios")
    else:
        print(f"⚠️  VERIFICACIÓN FALLÓ")
        print(f"❌ Hay archivos modificados o faltantes")
        print(f"💡 Revisar cambios o crear nuevo snapshot si es correcto")
    print(f"{'='*60}")

    return all_ok


def main():
    parser = argparse.ArgumentParser(
        description='Validación de integridad para archivos financieros audit-ready',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:

  # Crear snapshot de integridad
  python scripts/audit_integrity_check.py

  # Verificar contra snapshot existente
  python scripts/audit_integrity_check.py --verify audit_checksums.json

  # Especificar directorio personalizado
  python scripts/audit_integrity_check.py --data-dir ./mi_carpeta_datos

Workflow recomendado:
  1. Crear snapshot antes de cambios importantes
  2. Hacer modificaciones en archivos
  3. Verificar integridad con --verify
  4. Si OK, hacer commit del nuevo snapshot

  git add audit_checksums.json
  git commit -m "audit: snapshot integridad $(date +%Y-%m-%d)"
        """
    )

    parser.add_argument(
        '--verify',
        metavar='SNAPSHOT_FILE',
        help='Verificar integridad contra snapshot existente'
    )

    parser.add_argument(
        '--data-dir',
        default='data',
        help='Directorio con archivos a auditar (default: data/)'
    )

    parser.add_argument(
        '--output',
        default='audit_checksums.json',
        help='Archivo de salida para snapshot (default: audit_checksums.json)'
    )

    args = parser.parse_args()

    if args.verify:
        # Modo verificación
        success = verify_snapshot(args.verify, args.data_dir)
        sys.exit(0 if success else 1)
    else:
        # Modo creación de snapshot
        create_snapshot(args.data_dir, args.output)
        sys.exit(0)


if __name__ == '__main__':
    main()
