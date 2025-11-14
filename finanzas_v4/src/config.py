# -*- coding: utf-8 -*-
"""
🔧 CONFIGURACIÓN DEL SISTEMA DE FINANZAS v4.0
===============================================

Define la estructura completa del archivo Excel y configuraciones del sistema.
Este archivo es la FUENTE ÚNICA DE VERDAD para la estructura de datos.

LECCIÓN APRENDIDA v3.0:
- ❌ Usar índices hardcodeados (ws.cell(row, 5, valor))
- ✅ Usar diccionario de estructura (ws.cell(row, ESTRUCTURA['campo']['col'], valor))
"""

from datetime import datetime

# ==============================================================================
# INFORMACIÓN DEL SISTEMA
# ==============================================================================

VERSION = "4.0"
FECHA_VERSION = "2025-11-14"
AUTOR = "Sistema Finanzas v4.0"

# ==============================================================================
# ESTRUCTURA DE HOJA: TRANSACCIONES
# ==============================================================================

ESTRUCTURA_TRANSACCIONES = {
    'Fecha': {
        'col': 1,
        'col_letra': 'A',
        'tipo': 'datetime',
        'formato': 'DD/MM/YYYY',
        'formato_openpyxl': 'DD/MM/YYYY',
        'requerido': True,
        'descripcion': 'Fecha de la transacción'
    },
    'Tipo': {
        'col': 2,
        'col_letra': 'B',
        'tipo': 'string',
        'valores_validos': ['INGRESO', 'EGRESO', 'TRANSFERENCIA'],
        'requerido': True,
        'descripcion': 'Tipo de transacción'
    },
    'Categoría': {
        'col': 3,
        'col_letra': 'C',
        'tipo': 'string',
        'valores_validos': ['Operaciones', 'Proyectos', 'Administrativo', 'Ventas', 'Otros'],
        'requerido': True,
        'descripcion': 'Categoría principal'
    },
    'Subcategoría': {
        'col': 4,
        'col_letra': 'D',
        'tipo': 'string',
        'requerido': False,
        'descripcion': 'Subcategoría detallada'
    },
    'Descripción': {
        'col': 5,
        'col_letra': 'E',
        'tipo': 'string',
        'requerido': True,
        'descripcion': 'Descripción de la transacción'
    },
    'Entidad': {
        'col': 6,
        'col_letra': 'F',
        'tipo': 'string',
        'requerido': False,
        'descripcion': 'Proveedor, cliente o entidad involucrada'
    },
    'Cuenta': {
        'col': 7,
        'col_letra': 'G',
        'tipo': 'string',
        'requerido': True,
        'descripcion': 'Cuenta bancaria o método'
    },
    'Método Pago': {
        'col': 8,
        'col_letra': 'H',
        'tipo': 'string',
        'valores_validos': ['Transferencia', 'Efectivo', 'Tarjeta', 'Cheque', 'SINPE', 'Otro'],
        'requerido': False,
        'descripcion': 'Método de pago utilizado'
    },
    'Monto': {
        'col': 9,
        'col_letra': 'I',
        'tipo': 'float',
        'formato': '₡#,##0.00',
        'formato_openpyxl': '₡#,##0.00',
        'requerido': True,
        'descripcion': 'Monto de la transacción'
    },
    'Referencia': {
        'col': 10,
        'col_letra': 'J',
        'tipo': 'string',
        'requerido': False,
        'descripcion': 'Número de referencia o comprobante'
    },
    'Notas': {
        'col': 11,
        'col_letra': 'K',
        'tipo': 'string',
        'requerido': False,
        'descripcion': 'Notas adicionales'
    },
    'Estado': {
        'col': 12,
        'col_letra': 'L',
        'tipo': 'string',
        'valores_validos': ['PAGADO', 'PENDIENTE', 'POR COBRAR', 'CANCELADO', 'PARCIAL'],
        'requerido': True,
        'descripcion': 'Estado de la transacción (CRÍTICO para CxP/CxC)'
    },
    'IVA': {
        'col': 13,
        'col_letra': 'M',
        'tipo': 'float',
        'formato': '₡#,##0.00',
        'formato_openpyxl': '₡#,##0.00',
        'requerido': False,
        'descripcion': 'Monto de IVA (13%)'
    },
    'Fecha Vencimiento': {
        'col': 14,
        'col_letra': 'N',
        'tipo': 'datetime',
        'formato': 'DD/MM/YYYY',
        'formato_openpyxl': 'DD/MM/YYYY',
        'requerido': False,
        'descripcion': 'Fecha de vencimiento para CxP/CxC'
    },
    'Número Factura': {
        'col': 15,
        'col_letra': 'O',
        'tipo': 'string',
        'requerido': False,
        'descripcion': 'Número de factura'
    }
}

# ==============================================================================
# ESTRUCTURA DE HOJA: CxP (CUENTAS POR PAGAR)
# ==============================================================================

ESTRUCTURA_CXP = {
    'Fecha': {'col': 1, 'col_letra': 'A'},
    'Descripción': {'col': 2, 'col_letra': 'B'},
    'Entidad': {'col': 3, 'col_letra': 'C'},
    'Monto': {'col': 4, 'col_letra': 'D', 'formato': '₡#,##0.00'},
    'Fecha Vencimiento': {'col': 5, 'col_letra': 'E'},
    'Días Pendientes': {'col': 6, 'col_letra': 'F'},
    'Estado': {'col': 7, 'col_letra': 'G'},
    'Notas': {'col': 8, 'col_letra': 'H'}
}

# ==============================================================================
# ESTRUCTURA DE HOJA: CxC (CUENTAS POR COBRAR)
# ==============================================================================

ESTRUCTURA_CXC = {
    'Fecha': {'col': 1, 'col_letra': 'A'},
    'Descripción': {'col': 2, 'col_letra': 'B'},
    'Cliente': {'col': 3, 'col_letra': 'C'},
    'Monto': {'col': 4, 'col_letra': 'D', 'formato': '₡#,##0.00'},
    'Fecha Vencimiento': {'col': 5, 'col_letra': 'E'},
    'Días Pendientes': {'col': 6, 'col_letra': 'F'},
    'Estado': {'col': 7, 'col_letra': 'G'},
    'Notas': {'col': 8, 'col_letra': 'H'}
}

# ==============================================================================
# CONFIGURACIÓN DE HOJAS
# ==============================================================================

HOJAS_CONFIG = {
    'TRANSACCIONES': {
        'estructura': ESTRUCTURA_TRANSACCIONES,
        'fila_inicio_datos': 2,
        'color_header': 'FF4472C4',  # Azul
        'protegida': False,
        'descripcion': 'Registro completo de transacciones'
    },
    'CxP': {
        'estructura': ESTRUCTURA_CXP,
        'fila_inicio_datos': 6,  # Empieza en fila 6 (headers + resumen)
        'color_header': 'FFFF0000',  # Rojo
        'protegida': True,  # Solo fórmulas, no editar manualmente
        'descripcion': 'Cuentas por pagar (Estado = PENDIENTE)'
    },
    'CxC': {
        'estructura': ESTRUCTURA_CXC,
        'fila_inicio_datos': 6,
        'color_header': 'FF00B050',  # Verde
        'protegida': True,
        'descripcion': 'Cuentas por cobrar (Estado = POR COBRAR)'
    },
    'RESUMEN': {
        'estructura': {},  # Dashboard, estructura libre
        'fila_inicio_datos': 1,
        'color_header': 'FF7030A0',  # Morado
        'protegida': False,
        'descripcion': 'Dashboard con métricas y gráficos'
    }
}

# ==============================================================================
# FÓRMULAS PARA CxP
# ==============================================================================

def get_formula_cxp(fila_cxp, max_fila_trans):
    """
    Genera fórmulas para CxP que buscan en TRANSACCIONES

    LECCIÓN v3.0: Las fórmulas deben buscar en columna CORRECTA
    - Estado está en columna L (col 12)
    - Usar referencias absolutas $L$2:$L$1000
    """
    formulas = {
        'Fecha': f'=IFERROR(INDEX(TRANSACCIONES!$A$2:$A${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="PENDIENTE",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxp-5})),"")',

        'Descripción': f'=IFERROR(INDEX(TRANSACCIONES!$E$2:$E${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="PENDIENTE",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxp-5})),"")',

        'Entidad': f'=IFERROR(INDEX(TRANSACCIONES!$F$2:$F${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="PENDIENTE",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxp-5})),"")',

        'Monto': f'=IFERROR(INDEX(TRANSACCIONES!$I$2:$I${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="PENDIENTE",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxp-5})),"")',

        'Fecha Vencimiento': f'=IFERROR(INDEX(TRANSACCIONES!$N$2:$N${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="PENDIENTE",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxp-5})),"")',

        'Días Pendientes': f'=IF(A{fila_cxp}="","",TODAY()-A{fila_cxp})',

        'Estado': f'=IF(A{fila_cxp}="","",IF(E{fila_cxp}<TODAY(),"VENCIDO",IF(E{fila_cxp}<=TODAY()+7,"POR VENCER","AL DÍA")))',

        'Notas': f'=IFERROR(INDEX(TRANSACCIONES!$K$2:$K${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="PENDIENTE",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxp-5})),"")'
    }
    return formulas

# ==============================================================================
# FÓRMULAS PARA CxC
# ==============================================================================

def get_formula_cxc(fila_cxc, max_fila_trans):
    """Genera fórmulas para CxC que buscan Estado = POR COBRAR"""
    formulas = {
        'Fecha': f'=IFERROR(INDEX(TRANSACCIONES!$A$2:$A${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="POR COBRAR",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxc-5})),"")',

        'Descripción': f'=IFERROR(INDEX(TRANSACCIONES!$E$2:$E${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="POR COBRAR",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxc-5})),"")',

        'Cliente': f'=IFERROR(INDEX(TRANSACCIONES!$F$2:$F${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="POR COBRAR",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxc-5})),"")',

        'Monto': f'=IFERROR(INDEX(TRANSACCIONES!$I$2:$I${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="POR COBRAR",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxc-5})),"")',

        'Fecha Vencimiento': f'=IFERROR(INDEX(TRANSACCIONES!$N$2:$N${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="POR COBRAR",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxc-5})),"")',

        'Días Pendientes': f'=IF(A{fila_cxc}="","",TODAY()-A{fila_cxc})',

        'Estado': f'=IF(A{fila_cxc}="","",IF(E{fila_cxc}<TODAY(),"VENCIDO",IF(E{fila_cxc}<=TODAY()+7,"POR VENCER","AL DÍA")))',

        'Notas': f'=IFERROR(INDEX(TRANSACCIONES!$K$2:$K${max_fila_trans},SMALL(IF(UPPER(TRANSACCIONES!$L$2:$L${max_fila_trans})="POR COBRAR",ROW(TRANSACCIONES!$L$2:$L${max_fila_trans})-ROW(TRANSACCIONES!$L$2)+1),{fila_cxc-5})),"")'
    }
    return formulas

# ==============================================================================
# VALIDACIONES
# ==============================================================================

VALIDACIONES = {
    'tipo_transaccion': ['INGRESO', 'EGRESO', 'TRANSFERENCIA'],
    'estado_transaccion': ['PAGADO', 'PENDIENTE', 'POR COBRAR', 'CANCELADO', 'PARCIAL'],
    'metodo_pago': ['Transferencia', 'Efectivo', 'Tarjeta', 'Cheque', 'SINPE', 'Otro'],
    'categorias': ['Operaciones', 'Proyectos', 'Administrativo', 'Ventas', 'Otros']
}

# ==============================================================================
# CONFIGURACIÓN DE RESPALDOS
# ==============================================================================

RESPALDOS_CONFIG = {
    'directorio': '../backups',
    'formato_nombre': 'v4.0_backup_{timestamp}.xlsx',
    'max_respaldos': 10,  # Mantener últimos 10 respaldos
    'auto_respaldo': True
}

# ==============================================================================
# CONFIGURACIÓN DE LOGS
# ==============================================================================

LOGS_CONFIG = {
    'directorio': '../logs',
    'nivel': 'INFO',
    'formato': '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    'archivo_auditoria': 'auditoria_{timestamp}.log',
    'archivo_migracion': 'migracion_{timestamp}.log'
}

# ==============================================================================
# FUNCIONES AUXILIARES
# ==============================================================================

def get_timestamp():
    """Retorna timestamp para nombres de archivo"""
    return datetime.now().strftime('%Y%m%d_%H%M%S')

def get_columna_por_nombre(nombre_campo, hoja='TRANSACCIONES'):
    """
    Retorna el índice de columna para un campo específico

    Uso:
        col = get_columna_por_nombre('Descripción')
        ws.cell(row, col, 'Mi descripción')
    """
    estructura = HOJAS_CONFIG[hoja]['estructura']
    if nombre_campo in estructura:
        return estructura[nombre_campo]['col']
    else:
        raise ValueError(f"Campo '{nombre_campo}' no existe en hoja '{hoja}'")

def get_formato_campo(nombre_campo, hoja='TRANSACCIONES'):
    """Retorna el formato de celda para un campo"""
    estructura = HOJAS_CONFIG[hoja]['estructura']
    if nombre_campo in estructura:
        return estructura[nombre_campo].get('formato_openpyxl', None)
    return None

def validar_valor_campo(nombre_campo, valor, hoja='TRANSACCIONES'):
    """
    Valida que un valor sea válido para un campo

    Returns:
        (bool, str): (es_valido, mensaje_error)
    """
    estructura = HOJAS_CONFIG[hoja]['estructura']

    if nombre_campo not in estructura:
        return False, f"Campo '{nombre_campo}' no existe"

    config = estructura[nombre_campo]

    # Validar requerido
    if config.get('requerido', False) and not valor:
        return False, f"Campo '{nombre_campo}' es requerido"

    # Validar valores válidos
    if 'valores_validos' in config:
        if valor not in config['valores_validos']:
            return False, f"Valor '{valor}' no válido para '{nombre_campo}'. Válidos: {config['valores_validos']}"

    # Validar tipo
    tipo = config.get('tipo')
    if tipo == 'float':
        try:
            float(valor)
        except (ValueError, TypeError):
            return False, f"Campo '{nombre_campo}' debe ser numérico"

    return True, ""

# ==============================================================================
# DATOS DE EJEMPLO
# ==============================================================================

TRANSACCIONES_EJEMPLO = [
    {
        'Fecha': '01/11/2025',
        'Tipo': 'EGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Nómina',
        'Descripción': 'Pago de salario - Noviembre 2025',
        'Entidad': 'Juan Pérez',
        'Cuenta': 'BAC Corriente',
        'Método Pago': 'Transferencia',
        'Monto': 500000,
        'Referencia': 'TRF-2025-001',
        'Notas': 'Pago quincenal',
        'Estado': 'PAGADO',
        'IVA': 0,
        'Fecha Vencimiento': '',
        'Número Factura': ''
    },
    {
        'Fecha': '05/11/2025',
        'Tipo': 'EGRESO',
        'Categoría': 'Operaciones',
        'Subcategoría': 'Proveedores',
        'Descripción': 'Compra de materiales',
        'Entidad': 'Proveedor XYZ',
        'Cuenta': 'BAC Corriente',
        'Método Pago': 'Transferencia',
        'Monto': 150000,
        'Referencia': 'FAC-001',
        'Notas': 'Materiales proyecto A',
        'Estado': 'PENDIENTE',
        'IVA': 19500,
        'Fecha Vencimiento': '20/11/2025',
        'Número Factura': 'FAC-2025-001'
    },
    {
        'Fecha': '10/11/2025',
        'Tipo': 'INGRESO',
        'Categoría': 'Ventas',
        'Subcategoría': 'Servicios',
        'Descripción': 'Pago cliente por servicio',
        'Entidad': 'Cliente ABC',
        'Cuenta': 'BAC Corriente',
        'Método Pago': 'SINPE',
        'Monto': 300000,
        'Referencia': 'SINPE-001',
        'Notas': 'Servicio de consultoría',
        'Estado': 'POR COBRAR',
        'IVA': 39000,
        'Fecha Vencimiento': '25/11/2025',
        'Número Factura': 'INV-2025-001'
    }
]

# ==============================================================================
# INFORMACIÓN DE VERSIÓN
# ==============================================================================

print(f"""
╔═══════════════════════════════════════════════════════════════════╗
║              CONFIGURACIÓN CARGADA - v{VERSION}                     ║
╠═══════════════════════════════════════════════════════════════════╣
║ ✅ Estructura de TRANSACCIONES: {len(ESTRUCTURA_TRANSACCIONES)} campos              ║
║ ✅ Estructura de CxP: {len(ESTRUCTURA_CXP)} campos                           ║
║ ✅ Estructura de CxC: {len(ESTRUCTURA_CXC)} campos                           ║
║ ✅ Validaciones configuradas: {len(VALIDACIONES)} tipos                    ║
║ ✅ Fecha versión: {FECHA_VERSION}                               ║
╚═══════════════════════════════════════════════════════════════════╝
""")
