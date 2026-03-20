# DHCP Search Tool v2.1

Herramienta interna de soporte para consultar, gestionar y crear reservas y ámbitos DHCP.  
Desarrollada en PowerShell 5.1 con interfaz WPF (dark theme).

---

## Requisitos previos

| Requisito | Detalle |
|---|---|
| PowerShell | Versión 5.1 o superior |
| RSAT – DHCP Server Tools | Necesario para los cmdlets `Get-DhcpServerv4*` / `Add-DhcpServerv4*` |
| Acceso de red | El equipo debe poder alcanzar el servidor DHCP (`dhcp-server` por defecto) |
| Permisos | El usuario que ejecute el script necesita permisos de administración en el servidor DHCP para crear o modificar reservas y ámbitos |

Para instalar las herramientas RSAT en Windows 10/11:
```powershell
Add-WindowsCapability -Online -Name 'Rsat.DHCP.Tools~~~~0.0.1.0'
```

---

## Ejecución

```powershell
# Directamente desde PowerShell
& ".\DHCP Search.ps1"

# O bien desde el menú Aplicacións de NRC_APP (LazyWinAdmin)
```

---

## Descripción general

La aplicación se compone de dos paneles principales:

```
┌─────────────────────────────────────────────────────────────────┐
│  Barra superior: título · selector de servidor · estado         │
├─────────────────┬───────────────────────────────────────────────┤
│  Panel izquierdo│  Panel derecho                                │
│  ─────────────  │  ──────────────────────────────────────────   │
│  Ámbitos DHCP   │  Barra de búsqueda global                     │
│  (lista)        │  Filtro rápido / contexto activo / + reserva  │
│                 │  DataGrid con resultados                       │
│  + Nuevo ámbito │  Barra de estado (conteo, errores)            │
└─────────────────┴───────────────────────────────────────────────┘
```

### Panel izquierdo – Ámbitos

- Lista todos los ámbitos DHCP disponibles en el servidor seleccionado.
- El botón **↻ Cargar ámbitos** (o la carga automática al arrancar) consulta el servidor y rellena la lista.
- El campo **Filtrar ámbitos...** permite buscar por ID de red o nombre en tiempo real, sin volver a consultar.
- Al seleccionar un ámbito se cargan automáticamente sus reservas y leases en el panel derecho.
- El botón **+ Nuevo ámbito** abre un formulario para crear un nuevo ámbito en el servidor.

### Panel derecho – Resultados

- La **barra de búsqueda global** permite buscar por IP, MAC, nombre de equipo o descripción en todos los ámbitos a la vez.  
  Si la caché de datos ya está cargada, la búsqueda es instantánea (local). Si no, se lanza una consulta directa al servidor.
- El **selector de tipo** (Todo / Reserva / Lease) filtra las filas del grid.
- El campo **Filtrar resultados visibles...** aplica un filtro adicional sobre lo ya cargado.
- El botón **+ Nueva reserva** aparece cuando hay un ámbito seleccionado y permite crear una reserva directamente en él.
- El botón **Limpiar** resetea búsqueda, filtros y selección.

### DataGrid

Cada fila representa una entrada del servidor DHCP con las siguientes columnas:

| Columna | Descripción |
|---|---|
| Tipo | `Reserva` o `Lease` |
| IP | Dirección IP asignada |
| MAC | Identificador de cliente (MAC address) |
| Nombre | Nombre del equipo |
| Ámbito | ScopeId (red del ámbito) |
| Estado | Estado del tipo (Active, Both, etc.) |
| Descripción | Descripción de la reserva |
| Expira | Fecha/hora de expiración del lease (`-` en reservas) |

---

## Gestión de reservas

### Crear una reserva nueva

1. Seleccionar el ámbito donde se quiere crear la reserva en el panel izquierdo.
2. Hacer clic en el botón **+ Nueva reserva** que aparece en la barra de contexto del panel derecho.
3. Rellenar el formulario:
   - **IP**: dirección dentro del rango del ámbito (ej. `10.95.4.50`)
   - **MAC**: dirección MAC del cliente, con guiones o dos puntos (ej. `00-10-2B-00-4D-00`)
   - **Nombre**: nombre descriptivo del equipo (opcional)
   - **Descripción**: texto libre (opcional)
4. Confirmar con **Guardar**. La reserva se creará en el servidor y el grid se recargará automáticamente.

### Editar una reserva existente

1. Asegurarse de estar en vista de ámbito (seleccionar un ámbito en el panel izquierdo) o localizar la reserva con una búsqueda global.
2. Seleccionar la fila de tipo `Reserva` en el DataGrid.
3. **Doble clic** sobre la fila para abrir el formulario de edición con los datos actuales.
4. Modificar los campos deseados (nombre, descripción; la IP y MAC también se pueden cambiar).
5. Confirmar con **Guardar**. La reserva antigua se elimina y se crea la nueva con los datos actualizados.

> **Nota:** Los leases no se pueden editar directamente. Al hacer doble clic en un lease, se copian sus datos al portapapeles.

---

## Crear un nuevo ámbito

1. Hacer clic en el botón **+ Nuevo ámbito** del panel izquierdo.
2. Rellenar el formulario:
   - **Red (ScopeId)**: dirección de red del ámbito (ej. `10.95.5.0`)
   - **Máscara**: máscara de subred (ej. `255.255.255.0`)
   - **Nombre**: nombre descriptivo del ámbito (ej. `Ejemplo - Planta 1`)
   - **Rango inicio**: primera IP del pool (ej. `10.95.5.10`)
   - **Rango fin**: última IP del pool (ej. `10.95.5.200`)
3. Confirmar con **Crear**. El ámbito se creará en el servidor y la lista de ámbitos se recargará.

---

## Atajos de teclado

| Atajo | Acción |
|---|---|
| `Enter` en el campo de búsqueda | Lanzar búsqueda global |
| `Ctrl+C` con fila seleccionada | Copiar datos de la fila (tabulado) al portapapeles |
| `Doble clic` en reserva | Abrir formulario de edición |
| `Doble clic` en lease | Copiar datos al portapapeles |

---

## Caché de datos

Al cargar la lista de ámbitos, la aplicación lanza en segundo plano una carga completa de todos los ámbitos (reservas + leases). Mientras esta carga está en curso, la barra de estado muestra `cargando datos en background...`. Una vez completada, las búsquedas globales son instantáneas porque se resuelven contra la caché local, sin volver a consultar el servidor.

---

## Servidor DHCP configurado

El servidor por defecto es `dhcp-server`. Para añadir más servidores, modificar la variable `$Script:Config.Servidores` al inicio del script:

```powershell
$Script:Config = @{
    ServidorPorDefecto = 'dhcp-server'
    Servidores         = @('dhcp-server', 'OTRO_SERVIDOR')
    ...
}
```

---

## Solución de problemas frecuentes

| Síntoma | Causa probable | Solución |
|---|---|---|
| Error al cargar ámbitos | Sin conectividad o sin permisos en el servidor | Verificar red y credenciales |
| Cmdlet no reconocido | RSAT no instalado | Instalar `Rsat.DHCP.Tools` |
| No aparecen reservas | El ámbito no tiene reservas creadas | Normal; solo se listan entradas existentes |
| Error al crear reserva | IP fuera del rango, MAC con formato incorrecto o IP ya en uso | Revisar los datos introducidos |
| Error al crear ámbito | ScopeId o rangos inválidos, o el ámbito ya existe | Revisar los datos introducidos |
