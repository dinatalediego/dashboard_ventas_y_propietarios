# Validaciones Marketing / CRM

Este paquete contiene scripts para atender 2 pedidos:

1. Auditar leads de formularios: validar si los leads de ALICANTO de mayo tienen correo y teléfono.
2. Generar BBDD negativa de exclusión: contactos NO perfil, no alcanza cuota o no respondió.

## Estructura

validaciones_marketing_crm/
├── data/
│   └── colocar_aqui_tu_archivo_leads.xlsx
├── outputs/
├── audit_leads_formularios.py
├── generar_bbdd_negativa.py
├── run_all_validaciones.ps1
├── requirements.txt
└── README.md

## Instalación

Desde PowerShell:

```powershell
cd validaciones_marketing_crm
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Uso rápido

1. Coloca tu export del CRM o formularios en la carpeta `data/`.
2. Renómbralo como:

```text
leads_crm.xlsx
```

3. Ejecuta:

```powershell
python audit_leads_formularios.py
python generar_bbdd_negativa.py
```

O ejecuta todo junto:

```powershell
.\run_all_validaciones.ps1
```

## Salidas

Se generan archivos en `outputs/`:

- `auditoria_leads_alicanto_mayo.xlsx`
- `bbdd_negativa_exclusion.xlsx`

## Columnas esperadas

Por defecto los scripts esperan estas columnas:

- proyecto
- fecha_registro
- correo
- telefono
- nombre
- documento
- estado_lead
- observacion

Si tus columnas tienen otros nombres, edita la sección `CONFIGURACIÓN` en cada script.
