#Requires -Version 5.1
#==============================================================================
#  DHCP Search Tool v2.0
#  Autor  : Sergio
#  Fecha  : 2025
#==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#==============================================================================
# MÓDULO: Ensamblados WPF
#==============================================================================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms   # solo para Clipboard

#==============================================================================
# MÓDULO: Configuración global
#==============================================================================
$Script:Config = @{
    ServidorPorDefecto = 'ASSCC3141S'
    Servidores         = @('ASSCC3141S')   # Ampliar según necesidad
    Version            = '2.1'
    TituloApp          = 'DHCP Search Tool'
}

#==============================================================================
# MÓDULO: Paleta de colores y estilos (dark theme)
#==============================================================================
$Script:Tema = @{
    # Fondos
    FondoApp        = '#0F1117'
    FondoPanel      = '#1A1D27'
    FondoTarjeta    = '#20242F'
    FondoInput      = '#2A2E3E'
    FondoBarra      = '#141720'
    FondoGrid       = '#181B25'
    FondoGridFila   = '#1E2130'
    FondoGridAlter  = '#222638'
    FondoGridSelect = '#2D3A5E'
    FondoGridHover  = '#252940'
    FondoBadge      = '#252B3D'
    FondoBadgeScope = '#1A2A3A'

    # Texto
    TextoPrimario   = '#E8EAF0'
    TextoSecundario = '#8B90A8'
    TextoTerciario  = '#555A70'
    TextoAccent     = '#5B8CFF'
    TextoVerde      = '#4ADE80'
    TextoRojo       = '#F87171'
    TextoAmbar      = '#FBBF24'
    TextoCian       = '#22D3EE'
    TextoBlanco     = '#FFFFFF'

    # Acentos
    Accent          = '#5B8CFF'
    AccentHover     = '#7BA3FF'
    AccentActivo    = '#3D6FFF'
    Verde           = '#22C55E'
    VerdeHover      = '#16A34A'
    Rojo            = '#EF4444'
    Ambar           = '#F59E0B'
    Cian            = '#06B6D4'

    # Bordes
    Borde           = '#2E3347'
    BordeAccent     = '#3D5A99'
    BordeFocus      = '#5B8CFF'

    # Cabecera grid
    CabeceraFondo   = '#252B3D'
    CabeceraTexto   = '#A0A8C0'

    # Estados
    EstadoInfo      = '#5B8CFF'
    EstadoOk        = '#22C55E'
    EstadoError     = '#EF4444'
    EstadoWarning   = '#F59E0B'
    EstadoCargando  = '#06B6D4'
}

#==============================================================================
# MÓDULO: Helpers de construcción WPF
#==============================================================================
function New-WpfColor {
    param([string]$Hex)
    [System.Windows.Media.ColorConverter]::ConvertFromString($Hex)
}

function New-WpfBrush {
    param([string]$Hex)
    New-Object System.Windows.Media.SolidColorBrush (New-WpfColor $Hex)
}

function New-WpfFont {
    param(
        [double]$Size = 13,
        [string]$Weight = 'Normal',
        [string]$Family = 'Segoe UI'
    )
    $f = New-Object System.Windows.Controls.TextBlock
    $fam = New-Object System.Windows.Media.FontFamily($Family)
    $w = [System.Windows.FontWeights]::$Weight
    return @{ Size = $Size; Family = $fam; Weight = $w }
}

function Set-WpfFont {
    param($Control, [double]$Size = 13, [string]$Weight = 'Normal', [string]$Family = 'Segoe UI')
    $Control.FontFamily = New-Object System.Windows.Media.FontFamily($Family)
    $Control.FontSize   = $Size
    $Control.FontWeight = [System.Windows.FontWeights]::$Weight
}

function New-WpfThickness {
    param($Left=0, $Top=0, $Right=0, $Bottom=0)
    New-Object System.Windows.Thickness($Left, $Top, $Right, $Bottom)
}

function New-WpfCornerRadius {
    param($R=0)
    New-Object System.Windows.CornerRadius($R)
}

function New-WpfGridLength {
    param($Value, [string]$Type = 'Pixel')
    switch ($Type) {
        'Auto'  { [System.Windows.GridLength]::Auto }
        'Star'  { New-Object System.Windows.GridLength($Value, [System.Windows.GridUnitType]::Star) }
        default { New-Object System.Windows.GridLength($Value, [System.Windows.GridUnitType]::Pixel) }
    }
}

function Add-GridColumn {
    param($Grid, $Width, [string]$Type = 'Pixel')
    $col = New-Object System.Windows.Controls.ColumnDefinition
    $col.Width = New-WpfGridLength $Width $Type
    $Grid.ColumnDefinitions.Add($col)
}

function Add-GridRow {
    param($Grid, $Height, [string]$Type = 'Pixel')
    $row = New-Object System.Windows.Controls.RowDefinition
    $row.Height = New-WpfGridLength $Height $Type
    $Grid.RowDefinitions.Add($row)
}

function Set-GridPos {
    param($Control, [int]$Col=0, [int]$Row=0, [int]$ColSpan=1, [int]$RowSpan=1)
    [System.Windows.Controls.Grid]::SetColumn($Control, $Col)
    [System.Windows.Controls.Grid]::SetRow($Control, $Row)
    if ($ColSpan -gt 1) { [System.Windows.Controls.Grid]::SetColumnSpan($Control, $ColSpan) }
    if ($RowSpan -gt 1) { [System.Windows.Controls.Grid]::SetRowSpan($Control, $RowSpan) }
}

function New-RoundedBorder {
    param(
        [string]$Background   = $Script:Tema.FondoTarjeta,
        [string]$BorderColor  = $Script:Tema.Borde,
        [double]$BorderWidth  = 1,
        [double]$CornerRadius = 8,
        [System.Windows.Thickness]$Padding = (New-WpfThickness 0 0 0 0)
    )
    $b = New-Object System.Windows.Controls.Border
    $b.Background       = New-WpfBrush $Background
    $b.BorderBrush      = New-WpfBrush $BorderColor
    $b.BorderThickness  = New-WpfThickness $BorderWidth $BorderWidth $BorderWidth $BorderWidth
    $b.CornerRadius     = New-WpfCornerRadius $CornerRadius
    $b.Padding          = $Padding
    return $b
}

function New-StyledButton {
    param(
        [string]$Text,
        [string]$BgColor   = $Script:Tema.Accent,
        [string]$FgColor   = $Script:Tema.TextoBlanco,
        [string]$HoverColor = $Script:Tema.AccentHover,
        [double]$Width     = [double]::NaN,
        [double]$Height    = 36,
        [double]$FontSize  = 13,
        [double]$Radius    = 6,
        [System.Windows.Thickness]$Padding = (New-WpfThickness 16 0 16 0)
    )
    $btn = New-Object System.Windows.Controls.Button
    $btn.Content         = $Text
    $btn.Height          = $Height
    $btn.Background      = New-WpfBrush $BgColor
    $btn.Foreground      = New-WpfBrush $FgColor
    $btn.BorderThickness = New-WpfThickness 0 0 0 0
    $btn.Padding         = $Padding
    $btn.Cursor          = [System.Windows.Input.Cursors]::Hand
    $btn.VerticalContentAlignment = 'Center'
    if ($Width -ne [double]::NaN) { $btn.Width = $Width }
    Set-WpfFont $btn $FontSize 'SemiBold'

    # Template con CornerRadius (requiere ControlTemplate)
    $template = New-Object System.Windows.Controls.ControlTemplate([System.Windows.Controls.Button])
    $factory   = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.Border])
    $factory.SetValue([System.Windows.Controls.Border]::BackgroundProperty,
        (New-Object System.Windows.TemplateBindingExtension([System.Windows.Controls.Control]::BackgroundProperty)))
    $factory.SetValue([System.Windows.Controls.Border]::CornerRadiusProperty,
        (New-WpfCornerRadius $Radius))
    $factory.SetValue([System.Windows.Controls.Border]::PaddingProperty,
        (New-Object System.Windows.TemplateBindingExtension([System.Windows.Controls.Control]::PaddingProperty)))

    $cpFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.ContentPresenter])
    $cpFactory.SetValue([System.Windows.Controls.ContentPresenter]::HorizontalAlignmentProperty,
        [System.Windows.HorizontalAlignment]::Center)
    $cpFactory.SetValue([System.Windows.Controls.ContentPresenter]::VerticalAlignmentProperty,
        [System.Windows.VerticalAlignment]::Center)
    $factory.AppendChild($cpFactory)
    $template.VisualTree = $factory
    $btn.Template = $template

    # Guardar colores en Tag para evitar problemas de closure en event handlers WPF
    $btn.Tag = [PSCustomObject]@{
        Normal = New-WpfBrush $BgColor
        Hover  = New-WpfBrush $HoverColor
        Press  = New-WpfBrush $Script:Tema.AccentActivo
    }

    $btn.Add_MouseEnter({ $this.Background = $this.Tag.Hover })
    $btn.Add_MouseLeave({ $this.Background = $this.Tag.Normal })
    $btn.Add_PreviewMouseLeftButtonDown({ $this.Background = $this.Tag.Press })
    $btn.Add_PreviewMouseLeftButtonUp({ $this.Background = $this.Tag.Hover })

    return $btn
}

function New-StyledTextBox {
    param(
        [string]$Placeholder = '',
        [double]$Height = 36,
        [double]$FontSize = 13
    )
    $tb = New-Object System.Windows.Controls.TextBox
    $tb.Height           = $Height
    $tb.Background       = New-WpfBrush $Script:Tema.FondoInput
    $tb.Foreground       = New-WpfBrush $Script:Tema.TextoPrimario
    $tb.CaretBrush       = New-WpfBrush $Script:Tema.Accent
    $tb.BorderBrush      = New-WpfBrush $Script:Tema.Borde
    $tb.BorderThickness  = New-WpfThickness 1 1 1 1
    $tb.Padding          = New-WpfThickness 10 0 10 0
    $tb.VerticalContentAlignment = 'Center'
    $tb.Tag              = $Placeholder
    Set-WpfFont $tb $FontSize

    $tb.Add_GotFocus({
        $this.BorderBrush = New-WpfBrush $Script:Tema.BordeFocus
        if ($this.Text -eq $this.Tag) {
            $this.Text = ''
            $this.Foreground = New-WpfBrush $Script:Tema.TextoPrimario
        }
    })
    $tb.Add_LostFocus({
        $this.BorderBrush = New-WpfBrush $Script:Tema.Borde
        if ([string]::IsNullOrWhiteSpace($this.Text)) {
            $this.Text = $this.Tag
            $this.Foreground = New-WpfBrush $Script:Tema.TextoTerciario
        }
    })

    if ($Placeholder) {
        $tb.Text       = $Placeholder
        $tb.Foreground = New-WpfBrush $Script:Tema.TextoTerciario
    }

    return $tb
}

function New-Label {
    param(
        [string]$Text,
        [string]$Color = $Script:Tema.TextoPrimario,
        [double]$Size  = 13,
        [string]$Weight = 'Normal'
    )
    $lbl = New-Object System.Windows.Controls.TextBlock
    $lbl.Text       = $Text
    $lbl.Foreground = New-WpfBrush $Color
    Set-WpfFont $lbl $Size $Weight
    return $lbl
}

function New-CustomDropdown {
    # Reemplaza ComboBox con un control custom totalmente dark:
    # Border visible + Popup con ListBox que controlamos por completo.
    param(
        [string[]]$Items,
        [int]$SelectedIndex = 0,
        [double]$Width  = 130,
        [double]$Height = 36
    )

    $bgBrush     = New-WpfBrush $Script:Tema.FondoInput
    $fgBrush     = New-WpfBrush $Script:Tema.TextoPrimario
    $borderBrush = New-WpfBrush $Script:Tema.Borde
    $popupBg     = New-WpfBrush $Script:Tema.FondoPanel

    # Estado interno
    $state = [PSCustomObject]@{
        SelectedItem      = $Items[$SelectedIndex]
        SelectedIndex     = $SelectedIndex
        Items             = $Items
        Handlers          = [System.Collections.Generic.List[scriptblock]]::new()
    }

    # ── Contenedor principal ──────────────────────────────────────
    $outerBorder = New-Object System.Windows.Controls.Border
    $outerBorder.Width           = $Width
    $outerBorder.Height          = $Height
    $outerBorder.Background      = $bgBrush
    $outerBorder.BorderBrush     = $borderBrush
    $outerBorder.BorderThickness = New-WpfThickness 1 1 1 1
    $outerBorder.CornerRadius    = New-WpfCornerRadius 4
    $outerBorder.Cursor          = [System.Windows.Input.Cursors]::Hand

    $innerGrid = New-Object System.Windows.Controls.Grid
    Add-GridColumn $innerGrid 1  'Star'
    Add-GridColumn $innerGrid 24 'Pixel'

    $lblSelected = New-Object System.Windows.Controls.TextBlock
    $lblSelected.Text               = $state.SelectedItem
    $lblSelected.Foreground         = $fgBrush
    $lblSelected.VerticalAlignment  = 'Center'
    $lblSelected.Margin             = New-WpfThickness 10 0 0 0
    $lblSelected.TextTrimming       = 'CharacterEllipsis'
    Set-WpfFont $lblSelected 12
    Set-GridPos $lblSelected 0 0

    $lblArrow = New-Object System.Windows.Controls.TextBlock
    $lblArrow.Text              = '▾'
    $lblArrow.Foreground        = New-WpfBrush $Script:Tema.TextoSecundario
    $lblArrow.FontSize          = 10
    $lblArrow.VerticalAlignment = 'Center'
    $lblArrow.HorizontalAlignment = 'Center'
    Set-GridPos $lblArrow 1 0

    $innerGrid.Children.Add($lblSelected) | Out-Null
    $innerGrid.Children.Add($lblArrow)    | Out-Null
    $outerBorder.Child = $innerGrid

    # ── Popup ─────────────────────────────────────────────────────
    $popup = New-Object System.Windows.Controls.Primitives.Popup
    $popup.PlacementTarget   = $outerBorder
    $popup.Placement         = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
    $popup.StaysOpen         = $false
    $popup.AllowsTransparency = $true
    $popup.Width             = $Width

    $popupBorder = New-Object System.Windows.Controls.Border
    $popupBorder.Background      = $popupBg
    $popupBorder.BorderBrush     = $borderBrush
    $popupBorder.BorderThickness = New-WpfThickness 1 1 1 1
    $popupBorder.CornerRadius    = New-WpfCornerRadius 4

    $listBox = New-Object System.Windows.Controls.ListBox
    $listBox.Background      = New-WpfBrush 'Transparent'
    $listBox.BorderThickness = New-WpfThickness 0 0 0 0
    $listBox.Foreground      = $fgBrush
    $listBox.Padding         = New-WpfThickness 0 4 0 4
    Set-WpfFont $listBox 12

    # Estilo items
    $iStyle = New-Object System.Windows.Style([System.Windows.Controls.ListBoxItem])
    $iStyle.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BackgroundProperty,
        (New-WpfBrush 'Transparent'))))
    $iStyle.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::ForegroundProperty, $fgBrush)))
    $iStyle.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::PaddingProperty, (New-WpfThickness 12 7 12 7))))
    $iStyle.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BorderThicknessProperty, (New-WpfThickness 0 0 0 0))))
    $iStyle.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.FrameworkElement]::CursorProperty, [System.Windows.Input.Cursors]::Hand)))

    $tH = New-Object System.Windows.Trigger
    $tH.Property = [System.Windows.UIElement]::IsMouseOverProperty
    $tH.Value    = $true
    $tH.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BackgroundProperty, (New-WpfBrush $Script:Tema.FondoBadge))))
    $tH.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::ForegroundProperty, (New-WpfBrush $Script:Tema.TextoAccent))))
    $iStyle.Triggers.Add($tH)

    $tS = New-Object System.Windows.Trigger
    $tS.Property = [System.Windows.Controls.ListBoxItem]::IsSelectedProperty
    $tS.Value    = $true
    $tS.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BackgroundProperty, (New-WpfBrush $Script:Tema.FondoBadge))))
    $tS.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::ForegroundProperty, (New-WpfBrush $Script:Tema.Accent))))
    $iStyle.Triggers.Add($tS)
    $listBox.ItemContainerStyle = $iStyle

    foreach ($it in $Items) { $listBox.Items.Add($it) | Out-Null }
    $listBox.SelectedIndex = $SelectedIndex

    $popupBorder.Child = $listBox
    $popup.Child       = $popupBorder

    # ── Eventos ───────────────────────────────────────────────────
    # Abrir popup al hacer clic en el border
    $outerBorder.Tag = [PSCustomObject]@{ Popup = $popup; State = $state; Label = $lblSelected; Border = $outerBorder }
    $outerBorder.Add_MouseLeftButtonUp({
        $ctx = $this.Tag
        $ctx.Popup.IsOpen = -not $ctx.Popup.IsOpen
        if ($ctx.Popup.IsOpen) {
            $this.BorderBrush = New-WpfBrush $Script:Tema.BordeFocus
        } else {
            $this.BorderBrush = New-WpfBrush $Script:Tema.Borde
        }
    })

    # Selección en el listbox
    $listBox.Tag = [PSCustomObject]@{ Popup = $popup; State = $state; Label = $lblSelected; Border = $outerBorder }
    $listBox.Add_SelectionChanged({
        $ctx = $this.Tag
        if ($null -ne $this.SelectedItem) {
            $ctx.State.SelectedItem  = $this.SelectedItem.ToString()
            $ctx.State.SelectedIndex = $this.SelectedIndex
            $ctx.Label.Text          = $ctx.State.SelectedItem
            $ctx.Popup.IsOpen        = $false
            $ctx.Border.BorderBrush  = New-WpfBrush $Script:Tema.Borde
            foreach ($h in $ctx.State.Handlers) { & $h }
        }
    })

    # Método Add_SelectionChanged para compatibilidad
    $state | Add-Member -MemberType ScriptMethod -Name 'Add_SelectionChanged' -Value {
        param([scriptblock]$Handler)
        $this.Handlers.Add($Handler)
    }
    $state | Add-Member -MemberType ScriptMethod -Name 'get_SelectedItem' -Value {
        return $this.SelectedItem
    }

    # Adjuntar popup a la ventana al cargarse
    $outerBorder.Add_Loaded({
        $win = [System.Windows.Window]::GetWindow($this)
        if ($null -ne $win) {
            $existingContent = $win.Content
            # Agregar popup al árbol visual si no está ya
        }
    })

    return [PSCustomObject]@{
        Control = $outerBorder
        Popup   = $popup
        State   = $state
        ListBox = $listBox
    }
}


function New-Separator {
    param([string]$Color = $Script:Tema.Borde, [string]$Orient = 'Horizontal')
    $sep = New-Object System.Windows.Controls.Separator
    $sep.Background = New-WpfBrush $Color
    if ($Orient -eq 'Vertical') {
        $sep.Width  = 1
        $sep.Margin = New-WpfThickness 0 4 0 4
    } else {
        $sep.Height = 1
        $sep.Margin = New-WpfThickness 0 4 0 4
    }
    return $sep
}

#==============================================================================
# MÓDULO: DataGrid estilizado
#==============================================================================
function New-StyledDataGrid {
    $dg = New-Object System.Windows.Controls.DataGrid
    $dg.AutoGenerateColumns         = $false
    $dg.IsReadOnly                  = $true
    $dg.SelectionMode               = 'Single'
    $dg.SelectionUnit               = 'FullRow'
    $dg.CanUserAddRows              = $false
    $dg.CanUserDeleteRows           = $false
    $dg.CanUserReorderColumns       = $true
    $dg.CanUserResizeColumns        = $true
    $dg.CanUserSortColumns          = $true
    $dg.GridLinesVisibility         = 'None'
    $dg.HeadersVisibility           = 'Column'
    $dg.RowDetailsVisibilityMode    = 'Collapsed'
    $dg.Background                  = New-WpfBrush $Script:Tema.FondoGrid
    $dg.Foreground                  = New-WpfBrush $Script:Tema.TextoPrimario
    $dg.BorderThickness             = New-WpfThickness 0 0 0 0
    $dg.RowBackground               = New-WpfBrush $Script:Tema.FondoGridFila
    $dg.AlternatingRowBackground    = New-WpfBrush $Script:Tema.FondoGridAlter
    $dg.HorizontalScrollBarVisibility = 'Auto'
    $dg.VerticalScrollBarVisibility   = 'Auto'
    Set-WpfFont $dg 12.5

    # Estilo de cabeceras
    $hdrStyle = New-Object System.Windows.Style([System.Windows.Controls.Primitives.DataGridColumnHeader])
    $setters = @(
        @{ P = [System.Windows.Controls.Control]::BackgroundProperty;   V = (New-WpfBrush $Script:Tema.CabeceraFondo) },
        @{ P = [System.Windows.Controls.Control]::ForegroundProperty;   V = (New-WpfBrush $Script:Tema.CabeceraTexto) },
        @{ P = [System.Windows.Controls.Control]::FontWeightProperty;   V = [System.Windows.FontWeights]::SemiBold },
        @{ P = [System.Windows.Controls.Control]::FontSizeProperty;     V = [double]11.5 },
        @{ P = [System.Windows.Controls.Control]::BorderThicknessProperty; V = (New-WpfThickness 0 0 0 1) },
        @{ P = [System.Windows.Controls.Control]::BorderBrushProperty;  V = (New-WpfBrush $Script:Tema.Borde) },
        @{ P = [System.Windows.Controls.Control]::PaddingProperty;      V = (New-WpfThickness 12 10 12 10) },
        @{ P = [System.Windows.Controls.Control]::HeightProperty;       V = [double]38 }
    )
    foreach ($s in $setters) {
        $setter = New-Object System.Windows.Setter($s.P, $s.V)
        $hdrStyle.Setters.Add($setter)
    }
    $dg.ColumnHeaderStyle = $hdrStyle

    # Estilo de filas
    $rowStyle = New-Object System.Windows.Style([System.Windows.Controls.DataGridRow])
    $rowHeight = New-Object System.Windows.Setter([System.Windows.Controls.DataGridRow]::HeightProperty, [double]34)
    $rowStyle.Setters.Add($rowHeight)
    $rowCursor = New-Object System.Windows.Setter([System.Windows.FrameworkElement]::CursorProperty,
        [System.Windows.Input.Cursors]::Hand)
    $rowStyle.Setters.Add($rowCursor)

    # Trigger hover
    $trigger = New-Object System.Windows.Trigger
    $trigger.Property = [System.Windows.UIElement]::IsMouseOverProperty
    $trigger.Value    = $true
    $setter = New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BackgroundProperty,
        (New-WpfBrush $Script:Tema.FondoGridHover))
    $trigger.Setters.Add($setter)
    $rowStyle.Triggers.Add($trigger)

    # Trigger selección
    $triggerSel = New-Object System.Windows.Trigger
    $triggerSel.Property = [System.Windows.Controls.DataGridRow]::IsSelectedProperty
    $triggerSel.Value    = $true
    $setterSel = New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BackgroundProperty,
        (New-WpfBrush $Script:Tema.FondoGridSelect))
    $triggerSel.Setters.Add($setterSel)
    $setterBorder = New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BorderBrushProperty,
        (New-WpfBrush $Script:Tema.Accent))
    $triggerSel.Setters.Add($setterBorder)
    $rowStyle.Triggers.Add($triggerSel)

    $dg.RowStyle = $rowStyle

    # Estilo de celdas
    $cellStyle = New-Object System.Windows.Style([System.Windows.Controls.DataGridCell])
    $cellBorder = New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BorderThicknessProperty,
        (New-WpfThickness 0 0 0 0))
    $cellStyle.Setters.Add($cellBorder)
    $cellPad = New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::PaddingProperty,
        (New-WpfThickness 12 0 12 0))
    $cellStyle.Setters.Add($cellPad)

    $triggerFocus = New-Object System.Windows.Trigger
    $triggerFocus.Property = [System.Windows.UIElement]::IsFocusedProperty
    $triggerFocus.Value    = $true
    $sf = New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BorderThicknessProperty,
        (New-WpfThickness 0 0 0 0))
    $triggerFocus.Setters.Add($sf)
    $cellStyle.Triggers.Add($triggerFocus)
    $dg.CellStyle = $cellStyle

    return $dg
}

function Add-DgColumn {
    param($DataGrid, [string]$Header, [string]$Binding, [double]$MinWidth = 80, [bool]$Star = $false)
    $col = New-Object System.Windows.Controls.DataGridTextColumn
    $col.Header   = $Header
    $col.Binding  = New-Object System.Windows.Data.Binding($Binding)
    $col.MinWidth = $MinWidth
    if ($Star) {
        $col.Width = New-Object System.Windows.Controls.DataGridLength(1,
            [System.Windows.Controls.DataGridLengthUnitType]::Star)
    }
    $DataGrid.Columns.Add($col)
}

#==============================================================================
# MÓDULO: Spinner / indicador de carga
#==============================================================================
function New-Spinner {
    # Círculo animado usando ProgressBar circular simulado con Ellipse
    $canvas = New-Object System.Windows.Controls.Canvas
    $canvas.Width  = 32
    $canvas.Height = 32

    $dots = @()
    for ($i = 0; $i -lt 8; $i++) {
        $e = New-Object System.Windows.Shapes.Ellipse
        $e.Width  = 5
        $e.Height = 5
        $angle = $i * 45 * [Math]::PI / 180
        $x = 13.5 + 11 * [Math]::Sin($angle)
        $y = 13.5 - 11 * [Math]::Cos($angle)
        [System.Windows.Controls.Canvas]::SetLeft($e, $x)
        [System.Windows.Controls.Canvas]::SetTop($e, $y)
        $alpha = [byte](40 + ($i * 27))
        $e.Fill = New-Object System.Windows.Media.SolidColorBrush(
            [System.Windows.Media.Color]::FromArgb($alpha, 91, 140, 255))
        $canvas.Children.Add($e) | Out-Null
        $dots += $e
    }

    $Script:SpinnerDots  = $dots
    $Script:SpinnerIndex = 0
    $Script:SpinnerTimer = New-Object System.Windows.Threading.DispatcherTimer
    $Script:SpinnerTimer.Interval = [TimeSpan]::FromMilliseconds(80)
    $Script:SpinnerTimer.Add_Tick({
        $Script:SpinnerIndex = ($Script:SpinnerIndex + 1) % 8
        for ($i = 0; $i -lt 8; $i++) {
            $pos   = ($i - $Script:SpinnerIndex + 8) % 8
            $alpha = [byte](30 + $pos * 30)
            $Script:SpinnerDots[$i].Fill = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.Color]::FromArgb($alpha, 91, 140, 255))
        }
    })

    return $canvas
}

#==============================================================================
# MÓDULO: Funciones DHCP (ejecutadas en runspace)
#==============================================================================
$Script:DhcpFunctions = {
    function Get-DhcpScopes {
        param([string]$Servidor)
        $scopes = Get-DhcpServerv4Scope -ComputerName $Servidor -ErrorAction Stop
        return $scopes | Select-Object ScopeId, Name, SubnetMask, StartRange, EndRange, State,
            @{N='LeaseDuration';E={$_.LeaseDuration.TotalHours.ToString('0.0') + 'h'}}
    }

    function Get-DhcpReservations {
        param([string]$Servidor, [string]$ScopeId)
        $reservations = Get-DhcpServerv4Reservation -ComputerName $Servidor -ScopeId $ScopeId -ErrorAction Stop
        return $reservations | Select-Object ScopeId, IPAddress, ClientId, Name, Description, Type
    }

    function Search-DhcpAll {
        param([string]$Servidor, [string]$Filtro)
        $scopes  = Get-DhcpServerv4Scope -ComputerName $Servidor -ErrorAction Stop
        $results = [System.Collections.Generic.List[object]]::new()

        foreach ($scope in $scopes) {
            # Reservas
            try {
                $res = Get-DhcpServerv4Reservation -ComputerName $Servidor -ScopeId $scope.ScopeId -ErrorAction Stop
                foreach ($r in $res) {
                    if ($r.IPAddress  -like "*$Filtro*" -or
                        $r.ClientId  -like "*$Filtro*" -or
                        $r.Name      -like "*$Filtro*" -or
                        $r.Description -like "*$Filtro*") {
                        $results.Add([PSCustomObject]@{
                            Tipo        = 'Reserva'
                            ScopeId     = $scope.ScopeId.ToString()
                            ScopeName   = $scope.Name
                            IPAddress   = $r.IPAddress.ToString()
                            MAC         = $r.ClientId.ToString()
                            Nombre      = if ($r.Name) { $r.Name } else { '' }
                            Descripcion = if ($r.Description) { $r.Description } else { '' }
                            Estado      = $r.Type.ToString()
                        })
                    }
                }
            } catch {}

            # Leases activos
            try {
                $leases = Get-DhcpServerv4Lease -ComputerName $Servidor -ScopeId $scope.ScopeId -ErrorAction Stop
                foreach ($l in $leases) {
                    if ($l.IPAddress   -like "*$Filtro*" -or
                        $l.ClientId   -like "*$Filtro*" -or
                        $l.HostName   -like "*$Filtro*") {
                        $results.Add([PSCustomObject]@{
                            Tipo        = 'Lease'
                            ScopeId     = $scope.ScopeId.ToString()
                            ScopeName   = $scope.Name
                            IPAddress   = $l.IPAddress.ToString()
                            MAC         = if ($l.ClientId) { $l.ClientId.ToString() } else { '' }
                            Nombre      = if ($l.HostName) { $l.HostName } else { '' }
                            Descripcion = ''
                            Estado      = $l.AddressState.ToString()
                        })
                    }
                }
            } catch {}
        }
        return $results
    }

    function Get-DhcpScopeDetail {
        param([string]$Servidor, [string]$ScopeId)
        $results = [System.Collections.Generic.List[object]]::new()

        # Reservas del ámbito
        try {
            $res = Get-DhcpServerv4Reservation -ComputerName $Servidor -ScopeId $ScopeId -ErrorAction Stop
            foreach ($r in $res) {
                $results.Add([PSCustomObject]@{
                    Tipo        = 'Reserva'
                    ScopeId     = $ScopeId.ToString()
                    IPAddress   = $r.IPAddress.ToString()
                    MAC         = if ($r.ClientId) { $r.ClientId.ToString() } else { '' }
                    Nombre      = if ($r.Name) { $r.Name } else { '' }
                    Descripcion = if ($r.Description) { $r.Description } else { '' }
                    Estado      = $r.Type.ToString()
                    Expira      = ''
                })
            }
        } catch {}

        # Leases del ámbito
        try {
            $leases = Get-DhcpServerv4Lease -ComputerName $Servidor -ScopeId $ScopeId -ErrorAction Stop
            foreach ($l in $leases) {
                $results.Add([PSCustomObject]@{
                    Tipo        = 'Lease'
                    ScopeId     = $ScopeId.ToString()
                    IPAddress   = $l.IPAddress.ToString()
                    MAC         = if ($l.ClientId) { $l.ClientId.ToString() } else { '' }
                    Nombre      = if ($l.HostName) { $l.HostName } else { '' }
                    Descripcion = ''
                    Estado      = $l.AddressState.ToString()
                    Expira      = if ($l.LeaseExpiryTime) { $l.LeaseExpiryTime.ToString('dd/MM/yyyy HH:mm') } else { '-' }
                })
            }
        } catch {}

        return $results
    }

    function Load-DhcpFullCache {
        # Carga completa de todos los ámbitos: reservas + leases
        # Se ejecuta en background tras cargar la lista de ámbitos
        param([string]$Servidor, [string[]]$ScopeIds)
        $results = [System.Collections.Generic.List[object]]::new()

        foreach ($scopeId in $ScopeIds) {
            # Reservas
            try {
                $res = Get-DhcpServerv4Reservation -ComputerName $Servidor -ScopeId $scopeId -ErrorAction Stop
                foreach ($r in $res) {
                    $results.Add([PSCustomObject]@{
                        Tipo        = 'Reserva'
                        ScopeId     = $scopeId
                        ScopeName   = $scopeId
                        IPAddress   = $r.IPAddress.ToString()
                        MAC         = if ($r.ClientId) { $r.ClientId.ToString() } else { '' }
                        Nombre      = if ($r.Name) { $r.Name } else { '' }
                        Descripcion = if ($r.Description) { $r.Description } else { '' }
                        Estado      = $r.Type.ToString()
                        Expira      = ''
                    })
                }
            } catch {}

            # Leases
            try {
                $leases = Get-DhcpServerv4Lease -ComputerName $Servidor -ScopeId $scopeId -ErrorAction Stop
                foreach ($l in $leases) {
                    $results.Add([PSCustomObject]@{
                        Tipo        = 'Lease'
                        ScopeId     = $scopeId
                        ScopeName   = $scopeId
                        IPAddress   = $l.IPAddress.ToString()
                        MAC         = if ($l.ClientId) { $l.ClientId.ToString() } else { '' }
                        Nombre      = if ($l.HostName) { $l.HostName } else { '' }
                        Descripcion = ''
                        Estado      = $l.AddressState.ToString()
                        Expira      = if ($l.LeaseExpiryTime) { $l.LeaseExpiryTime.ToString('dd/MM/yyyy HH:mm') } else { '-' }
                    })
                }
            } catch {}
        }
        return $results
    }

    function New-DhcpReservation {
        param([string]$Servidor, [string]$ScopeId, [string]$IP, [string]$MAC, [string]$Nombre, [string]$Descripcion)
        Add-DhcpServerv4Reservation -ComputerName $Servidor -ScopeId $ScopeId -IPAddress $IP `
            -ClientId $MAC -Name $Nombre -Description $Descripcion -Type Both -ErrorAction Stop
    }

    function Edit-DhcpReservation {
        param([string]$Servidor, [string]$ScopeId, [string]$OldIP, [string]$IP, [string]$MAC, [string]$Nombre, [string]$Descripcion)
        Remove-DhcpServerv4Reservation -ComputerName $Servidor -ScopeId $ScopeId -IPAddress $OldIP -ErrorAction SilentlyContinue
        Add-DhcpServerv4Reservation -ComputerName $Servidor -ScopeId $ScopeId -IPAddress $IP `
            -ClientId $MAC -Name $Nombre -Description $Descripcion -Type Both -ErrorAction Stop
    }

    function New-DhcpScopeCustom {
        param([string]$Servidor, [string]$ScopeId, [string]$Mascara, [string]$Nombre, [string]$RangoInicio, [string]$RangoFin)
        Add-DhcpServerv4Scope -ComputerName $Servidor -Name $Nombre -SubnetMask $Mascara `
            -StartRange $RangoInicio -EndRange $RangoFin -State Active -ErrorAction Stop
    }
}

#==============================================================================
# MÓDULO: Motor de Runspaces asíncronos
#==============================================================================
function Invoke-AsyncDhcp {
    param(
        [scriptblock]$Trabajo,
        [hashtable]$Argumentos,
        [scriptblock]$OnComplete,
        [scriptblock]$OnError
    )

    $rsPool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 1)
    $rsPool.ApartmentState = 'STA'
    $rsPool.Open()

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.RunspacePool = $rsPool

    # Inyectar funciones DHCP en el runspace
    $ps.AddScript($Script:DhcpFunctions) | Out-Null
    $ps.AddScript($Trabajo) | Out-Null
    foreach ($kv in $Argumentos.GetEnumerator()) {
        $ps.AddParameter($kv.Key, $kv.Value) | Out-Null
    }

    $handle = $ps.BeginInvoke()

    # Encapsular todo el contexto en un objeto para evitar problemas de closure en PowerShell
    $ctx = [PSCustomObject]@{
        PS         = $ps
        Handle     = $handle
        Pool       = $rsPool
        OnComplete = $OnComplete
        OnError    = $OnError
    }

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(150)
    $timer.Tag = $ctx   # El Tag del timer lleva el contexto — sin closure

    $timer.Add_Tick({
        $c = $this.Tag   # recuperar contexto desde Tag
        if ($c.Handle.IsCompleted) {
            $this.Stop()
            try {
                $output = $c.PS.EndInvoke($c.Handle)
                $errors = $c.PS.Streams.Error
                $c.PS.Dispose()
                $c.Pool.Dispose()
                if ($errors.Count -gt 0 -and $null -ne $c.OnError) {
                    & $c.OnError $errors[0].Exception.Message
                } else {
                    & $c.OnComplete $output
                }
            } catch {
                try { $c.PS.Dispose()   } catch {}
                try { $c.Pool.Dispose() } catch {}
                if ($null -ne $c.OnError) { & $c.OnError $_.Exception.Message }
            }
        }
    })

    $timer.Start()
}

#==============================================================================
# MÓDULO: Construcción de la interfaz principal
#==============================================================================

# ── Ventana principal ──────────────────────────────────────────────────────────
$Script:Window = New-Object System.Windows.Window
$Script:Window.Title          = "$($Script:Config.TituloApp) v$($Script:Config.Version)"
$Script:Window.Width          = 1100
$Script:Window.Height         = 740
$Script:Window.MinWidth       = 900
$Script:Window.MinHeight      = 600
$Script:Window.WindowStartupLocation = 'CenterScreen'
$Script:Window.Background     = New-WpfBrush $Script:Tema.FondoApp
$Script:Window.AllowsTransparency = $false
$Script:Window.FontFamily     = New-Object System.Windows.Media.FontFamily('Segoe UI')

# ── Layout raíz ───────────────────────────────────────────────────────────────
$rootGrid = New-Object System.Windows.Controls.Grid
Add-GridRow $rootGrid 54 'Pixel'    # Barra superior
Add-GridRow $rootGrid 1  'Star'     # Contenido

# ── Barra superior ────────────────────────────────────────────────────────────
$barraFondo = New-Object System.Windows.Controls.Border
$barraFondo.Background      = New-WpfBrush $Script:Tema.FondoBarra
$barraFondo.BorderBrush     = New-WpfBrush $Script:Tema.Borde
$barraFondo.BorderThickness = New-WpfThickness 0 0 0 1
Set-GridPos $barraFondo 0 0

$barraGrid = New-Object System.Windows.Controls.Grid
Add-GridColumn $barraGrid 0 'Auto'    # Logo + título
Add-GridColumn $barraGrid 1 'Star'    # espacio
Add-GridColumn $barraGrid 0 'Auto'    # Servidor selector
Add-GridColumn $barraGrid 0 'Auto'    # Estado dot
$barraFondo.Child = $barraGrid

# Logo/título
$logoPanel = New-Object System.Windows.Controls.StackPanel
$logoPanel.Orientation = 'Horizontal'
$logoPanel.VerticalAlignment = 'Center'
$logoPanel.Margin = New-WpfThickness 20 0 0 0
Set-GridPos $logoPanel 0 0

$logoIcon = New-Label '⬡' $Script:Tema.Accent 20 'Bold'
$logoIcon.VerticalAlignment = 'Center'
$logoIcon.Margin = New-WpfThickness 0 0 8 0
$logoTitle = New-Label 'DHCP Search Tool' $Script:Tema.TextoPrimario 15 'SemiBold'
$logoTitle.VerticalAlignment = 'Center'
$logoVersion = New-Label " v$($Script:Config.Version)" $Script:Tema.TextoTerciario 11
$logoVersion.VerticalAlignment = 'Center'
$logoVersion.Margin = New-WpfThickness 4 2 0 0
$logoPanel.Children.Add($logoIcon)    | Out-Null
$logoPanel.Children.Add($logoTitle)   | Out-Null
$logoPanel.Children.Add($logoVersion) | Out-Null

# Selector de servidor
$servidorPanel = New-Object System.Windows.Controls.StackPanel
$servidorPanel.Orientation   = 'Horizontal'
$servidorPanel.VerticalAlignment = 'Center'
$servidorPanel.Margin        = New-WpfThickness 0 0 16 0
Set-GridPos $servidorPanel 2 0

$lblServidor = New-Label 'Servidor:' $Script:Tema.TextoSecundario 11
$lblServidor.VerticalAlignment = 'Center'
$lblServidor.Margin = New-WpfThickness 0 0 8 0

$Script:DdServidor = New-CustomDropdown -Items $Script:Config.Servidores -SelectedIndex 0 -Width 160 -Height 30
$Script:CmbServidor = $Script:DdServidor.State   # alias para compatibilidad de eventos

$servidorPanel.Children.Add($lblServidor)                  | Out-Null
$servidorPanel.Children.Add($Script:DdServidor.Control)    | Out-Null
$servidorPanel.Children.Add($Script:DdServidor.Popup)      | Out-Null

# Indicador de conexión
$Script:DotConexion = New-Object System.Windows.Shapes.Ellipse
$Script:DotConexion.Width  = 8
$Script:DotConexion.Height = 8
$Script:DotConexion.Fill   = New-WpfBrush $Script:Tema.TextoTerciario
$Script:DotConexion.VerticalAlignment = 'Center'
$Script:DotConexion.Margin = New-WpfThickness 0 0 20 0
Set-GridPos $Script:DotConexion 3 0

$barraGrid.Children.Add($logoPanel)            | Out-Null
$barraGrid.Children.Add($servidorPanel)        | Out-Null
$barraGrid.Children.Add($Script:DotConexion)   | Out-Null

# ── Área de contenido principal ───────────────────────────────────────────────
$contentBorder = New-Object System.Windows.Controls.Border
$contentBorder.Padding = New-WpfThickness 16 14 16 14
Set-GridPos $contentBorder 0 1

$contentGrid = New-Object System.Windows.Controls.Grid
Add-GridColumn $contentGrid 260 'Pixel'     # Panel izquierdo (ámbitos)
Add-GridColumn $contentGrid 10  'Pixel'     # Separador
Add-GridColumn $contentGrid 1   'Star'      # Panel derecho (resultados)
$contentBorder.Child = $contentGrid

# ── Panel izquierdo – Ámbitos ─────────────────────────────────────────────────
$leftPanel = New-Object System.Windows.Controls.Grid
Add-GridRow $leftPanel 0  'Auto'    # Fila 0: Título ámbitos
Add-GridRow $leftPanel 8  'Pixel'   # Fila 1: espacio
Add-GridRow $leftPanel 36 'Pixel'   # Fila 2: Botón recargar ámbitos
Add-GridRow $leftPanel 8  'Pixel'   # Fila 3: espacio
Add-GridRow $leftPanel 30 'Pixel'   # Fila 4: Buscador de ámbitos
Add-GridRow $leftPanel 8  'Pixel'   # Fila 5: espacio
Add-GridRow $leftPanel 1  'Star'    # Fila 6: Lista ámbitos
Add-GridRow $leftPanel 8  'Pixel'   # Fila 7: espacio
Add-GridRow $leftPanel 0  'Auto'    # Fila 8: Badge conteo
Add-GridRow $leftPanel 8  'Pixel'   # Fila 9: espacio
Add-GridRow $leftPanel 36 'Pixel'   # Fila 10: Botón nuevo ámbito
Set-GridPos $leftPanel 0 0

# Título panel izquierdo
$leftTitlePanel = New-Object System.Windows.Controls.StackPanel
$leftTitlePanel.Orientation = 'Horizontal'
Set-GridPos $leftTitlePanel 0 0

$leftTitle = New-Label 'Ámbitos DHCP' $Script:Tema.TextoPrimario 13 'SemiBold'
$leftTitle.VerticalAlignment = 'Center'
$Script:BadgeAmbitos = New-Label ' 0' $Script:Tema.TextoSecundario 11
$Script:BadgeAmbitos.VerticalAlignment = 'Center'
$Script:BadgeAmbitos.Margin = New-WpfThickness 6 1 0 0
$leftTitlePanel.Children.Add($leftTitle)         | Out-Null
$leftTitlePanel.Children.Add($Script:BadgeAmbitos) | Out-Null

# Botón recargar ámbitos
$Script:BtnRecargarAmbitos = New-StyledButton -Text '↻  Cargar ámbitos' `
    -BgColor $Script:Tema.FondoTarjeta -FgColor $Script:Tema.TextoPrimario `
    -HoverColor $Script:Tema.FondoInput -Height 32 -FontSize 12
$Script:BtnRecargarAmbitos.HorizontalAlignment = 'Stretch'
Set-GridPos $Script:BtnRecargarAmbitos 0 2

# Buscador de ámbitos
$Script:TxtFiltroAmbitos = New-StyledTextBox -Placeholder 'Filtrar ámbitos...' -Height 30 -FontSize 11.5
Set-GridPos $Script:TxtFiltroAmbitos 0 4

# Lista de ámbitos
$ambitosBorder = New-RoundedBorder $Script:Tema.FondoPanel $Script:Tema.Borde 1 8
Set-GridPos $ambitosBorder 0 6

$Script:ListaAmbitos = New-Object System.Windows.Controls.ListBox
$Script:ListaAmbitos.Background       = New-WpfBrush 'Transparent'
$Script:ListaAmbitos.BorderThickness  = New-WpfThickness 0 0 0 0
$Script:ListaAmbitos.Foreground       = New-WpfBrush $Script:Tema.TextoPrimario
[System.Windows.Controls.ScrollViewer]::SetHorizontalScrollBarVisibility(
    $Script:ListaAmbitos, [System.Windows.Controls.ScrollBarVisibility]::Disabled)
Set-WpfFont $Script:ListaAmbitos 12

# Estilo items de la lista
$itemStyle = New-Object System.Windows.Style([System.Windows.Controls.ListBoxItem])
$itemPad = New-Object System.Windows.Setter([System.Windows.Controls.Control]::PaddingProperty,
    (New-WpfThickness 10 8 10 8))
$itemStyle.Setters.Add($itemPad)
$itemBorder = New-Object System.Windows.Setter([System.Windows.Controls.Control]::BorderThicknessProperty,
    (New-WpfThickness 0 0 0 1))
$itemStyle.Setters.Add($itemBorder)
$itemBorderColor = New-Object System.Windows.Setter([System.Windows.Controls.Control]::BorderBrushProperty,
    (New-WpfBrush $Script:Tema.Borde))
$itemStyle.Setters.Add($itemBorderColor)
$itemCursor = New-Object System.Windows.Setter([System.Windows.FrameworkElement]::CursorProperty,
    [System.Windows.Input.Cursors]::Hand)
$itemStyle.Setters.Add($itemCursor)

$trigItemHover = New-Object System.Windows.Trigger
$trigItemHover.Property = [System.Windows.UIElement]::IsMouseOverProperty
$trigItemHover.Value    = $true
$trigItemHover.Setters.Add((New-Object System.Windows.Setter(
    [System.Windows.Controls.Control]::BackgroundProperty,
    (New-WpfBrush $Script:Tema.FondoInput))))
$itemStyle.Triggers.Add($trigItemHover)

$trigItemSel = New-Object System.Windows.Trigger
$trigItemSel.Property = [System.Windows.Controls.ListBoxItem]::IsSelectedProperty
$trigItemSel.Value    = $true
$trigItemSel.Setters.Add((New-Object System.Windows.Setter(
    [System.Windows.Controls.Control]::BackgroundProperty,
    (New-WpfBrush $Script:Tema.FondoBadge))))
$trigItemSel.Setters.Add((New-Object System.Windows.Setter(
    [System.Windows.Controls.Control]::ForegroundProperty,
    (New-WpfBrush $Script:Tema.Accent))))
$itemStyle.Triggers.Add($trigItemSel)
$Script:ListaAmbitos.ItemContainerStyle = $itemStyle

$ambitosBorder.Child = $Script:ListaAmbitos

# Badge conteo ámbito seleccionado
$Script:BadgeScopeInfo = New-Label '' $Script:Tema.TextoTerciario 10.5
$Script:BadgeScopeInfo.TextWrapping = 'Wrap'
Set-GridPos $Script:BadgeScopeInfo 0 8

# Botón nuevo ámbito
$Script:BtnNuevoAmbito = New-StyledButton -Text '+  Nuevo ámbito' `
    -BgColor $Script:Tema.Verde -FgColor $Script:Tema.TextoBlanco `
    -HoverColor $Script:Tema.VerdeHover -Height 32 -FontSize 12
$Script:BtnNuevoAmbito.HorizontalAlignment = 'Stretch'
Set-GridPos $Script:BtnNuevoAmbito 0 10

$leftPanel.Children.Add($leftTitlePanel)              | Out-Null
$leftPanel.Children.Add($Script:BtnRecargarAmbitos)   | Out-Null
$leftPanel.Children.Add($Script:TxtFiltroAmbitos)     | Out-Null
$leftPanel.Children.Add($ambitosBorder)               | Out-Null
$leftPanel.Children.Add($Script:BadgeScopeInfo)       | Out-Null
$leftPanel.Children.Add($Script:BtnNuevoAmbito)       | Out-Null

# ── Panel derecho ─────────────────────────────────────────────────────────────
$rightPanel = New-Object System.Windows.Controls.Grid
Add-GridRow $rightPanel 0  'Auto'   # Barra búsqueda
Add-GridRow $rightPanel 10 'Pixel'
Add-GridRow $rightPanel 0  'Auto'   # Barra filtro en ámbito / tabs
Add-GridRow $rightPanel 10 'Pixel'
Add-GridRow $rightPanel 1  'Star'   # Grid resultados
Add-GridRow $rightPanel 8  'Pixel'
Add-GridRow $rightPanel 28 'Pixel'  # Barra de estado
Set-GridPos $rightPanel 2 0

# ── Barra de búsqueda global ──────────────────────────────────────────────────
$searchCard = New-RoundedBorder $Script:Tema.FondoPanel $Script:Tema.Borde 1 8 (New-WpfThickness 14 0 14 0)
Set-GridPos $searchCard 0 0

$searchInnerGrid = New-Object System.Windows.Controls.Grid
Add-GridColumn $searchInnerGrid 1  'Star'
Add-GridColumn $searchInnerGrid 10 'Pixel'
Add-GridColumn $searchInnerGrid 0  'Auto'
Add-GridColumn $searchInnerGrid 8  'Pixel'
Add-GridColumn $searchInnerGrid 0  'Auto'
Add-GridColumn $searchInnerGrid 8  'Pixel'
Add-GridColumn $searchInnerGrid 0  'Auto'
$searchInnerGrid.Margin = New-WpfThickness 0 8 0 8

# TextBox búsqueda
$Script:TxtBuscar = New-StyledTextBox -Placeholder 'Buscar por IP, MAC, nombre de equipo o reserva...' -Height 38 -FontSize 13
Set-GridPos $Script:TxtBuscar 0 0

# Botón buscar
$Script:BtnBuscar = New-StyledButton -Text '⌕  Buscar' -Height 38 -Width 110
Set-GridPos $Script:BtnBuscar 2 0

# Tipo de búsqueda selector (dropdown custom)
$Script:DdTipoBusqueda = New-CustomDropdown -Items @('Todo','Reserva','Lease') -SelectedIndex 0 -Width 130 -Height 38
$Script:CmbTipoBusqueda = $Script:DdTipoBusqueda.State   # alias para compatibilidad
Set-GridPos $Script:DdTipoBusqueda.Control 4 0

# Botón limpiar
$Script:BtnLimpiar = New-StyledButton -Text 'Limpiar' `
    -BgColor $Script:Tema.FondoInput -FgColor $Script:Tema.TextoSecundario `
    -HoverColor $Script:Tema.Borde -Height 38 -Width 70 -FontSize 11
Set-GridPos $Script:BtnLimpiar 6 0

$searchInnerGrid.Children.Add($Script:TxtBuscar)          | Out-Null
$searchInnerGrid.Children.Add($Script:BtnBuscar)          | Out-Null
$searchInnerGrid.Children.Add($Script:DdTipoBusqueda.Control) | Out-Null
$searchInnerGrid.Children.Add($Script:DdTipoBusqueda.Popup)   | Out-Null
$searchInnerGrid.Children.Add($Script:BtnLimpiar)         | Out-Null
$searchCard.Child = $searchInnerGrid

# ── Barra de filtro de ámbito / contexto activo ───────────────────────────────
$filterBar = New-Object System.Windows.Controls.Grid
Add-GridColumn $filterBar 0 'Auto'
Add-GridColumn $filterBar 1 'Star'
Add-GridColumn $filterBar 0 'Auto'
Add-GridColumn $filterBar 8  'Pixel'
Add-GridColumn $filterBar 0  'Auto'
Set-GridPos $filterBar 0 2

$Script:LblContexto = New-Label 'Todos los ámbitos' $Script:Tema.TextoSecundario 11.5
$Script:LblContexto.VerticalAlignment = 'Center'
Set-GridPos $Script:LblContexto 0 0

$Script:TxtFiltroRapido = New-StyledTextBox -Placeholder 'Filtrar resultados visibles...' -Height 30 -FontSize 12
$Script:TxtFiltroRapido.MaxWidth = 240
$Script:TxtFiltroRapido.HorizontalAlignment = 'Right'
Set-GridPos $Script:TxtFiltroRapido 2 0

# Botón nueva reserva (visible solo con ámbito seleccionado)
$Script:BtnNuevaReserva = New-StyledButton -Text '+  Nueva reserva' `
    -BgColor $Script:Tema.Verde -FgColor $Script:Tema.TextoBlanco `
    -HoverColor $Script:Tema.VerdeHover -Height 30 -Width 140 -FontSize 11.5
$Script:BtnNuevaReserva.Visibility = 'Collapsed'
Set-GridPos $Script:BtnNuevaReserva 4 0

$filterBar.Children.Add($Script:LblContexto)       | Out-Null
$filterBar.Children.Add($Script:TxtFiltroRapido)   | Out-Null
$filterBar.Children.Add($Script:BtnNuevaReserva)   | Out-Null

# ── DataGrid de resultados ────────────────────────────────────────────────────
$gridWrapper = New-Object System.Windows.Controls.Border
$gridWrapper.Background     = New-WpfBrush $Script:Tema.FondoGrid
$gridWrapper.CornerRadius   = New-WpfCornerRadius 8
$gridWrapper.ClipToBounds   = $true
Set-GridPos $gridWrapper 0 4

$gridContainer = New-Object System.Windows.Controls.Grid
Add-GridRow $gridContainer 1 'Star'
Add-GridRow $gridContainer 0 'Auto'   # Panel vacío / spinner

$Script:DataGrid = New-StyledDataGrid
Set-GridPos $Script:DataGrid 0 0
$gridContainer.Children.Add($Script:DataGrid) | Out-Null

# Panel overlay (spinner + mensaje vacío)
$Script:PanelOverlay = New-Object System.Windows.Controls.Border
$Script:PanelOverlay.Background          = New-WpfBrush "$($Script:Tema.FondoGrid)CC"
$Script:PanelOverlay.HorizontalAlignment = 'Stretch'
$Script:PanelOverlay.VerticalAlignment   = 'Stretch'
$Script:PanelOverlay.Visibility          = 'Collapsed'
Set-GridPos $Script:PanelOverlay 0 0
[System.Windows.Controls.Grid]::SetRowSpan($Script:PanelOverlay, 2)

$overlayStack = New-Object System.Windows.Controls.StackPanel
$overlayStack.HorizontalAlignment = 'Center'
$overlayStack.VerticalAlignment   = 'Center'

$Script:SpinnerContainer = New-Object System.Windows.Controls.Border
$Script:SpinnerContainer.HorizontalAlignment = 'Center'
$spinnerCtrl = New-Spinner
$Script:SpinnerContainer.Child = $spinnerCtrl

$Script:LblOverlay = New-Label 'Consultando servidor DHCP...' $Script:Tema.TextoSecundario 13
$Script:LblOverlay.HorizontalAlignment = 'Center'
$Script:LblOverlay.Margin = New-WpfThickness 0 12 0 0

$overlayStack.Children.Add($Script:SpinnerContainer) | Out-Null
$overlayStack.Children.Add($Script:LblOverlay)       | Out-Null
$Script:PanelOverlay.Child = $overlayStack

$gridContainer.Children.Add($Script:PanelOverlay) | Out-Null
$gridWrapper.Child = $gridContainer

# ── Barra de estado inferior ──────────────────────────────────────────────────
$statusBar = New-Object System.Windows.Controls.Grid
Add-GridColumn $statusBar 8 'Pixel'
Add-GridColumn $statusBar 0 'Auto'
Add-GridColumn $statusBar 1 'Star'
Add-GridColumn $statusBar 0 'Auto'
Set-GridPos $statusBar 0 6

$Script:DotEstado = New-Object System.Windows.Shapes.Ellipse
$Script:DotEstado.Width  = 7
$Script:DotEstado.Height = 7
$Script:DotEstado.Fill   = New-WpfBrush $Script:Tema.EstadoInfo
$Script:DotEstado.VerticalAlignment = 'Center'
Set-GridPos $Script:DotEstado 1 0

$Script:LblEstado = New-Label 'Listo' $Script:Tema.TextoSecundario 11.5
$Script:LblEstado.VerticalAlignment = 'Center'
$Script:LblEstado.Margin = New-WpfThickness 7 0 0 0
Set-GridPos $Script:LblEstado 2 0

$Script:LblConteo = New-Label '' $Script:Tema.TextoTerciario 11
$Script:LblConteo.VerticalAlignment  = 'Center'
$Script:LblConteo.HorizontalAlignment = 'Right'
Set-GridPos $Script:LblConteo 3 0

$statusBar.Children.Add($Script:DotEstado)  | Out-Null
$statusBar.Children.Add($Script:LblEstado)  | Out-Null
$statusBar.Children.Add($Script:LblConteo)  | Out-Null

# Columnas del DataGrid
Add-DgColumn $Script:DataGrid 'Tipo'        'Tipo'        65
Add-DgColumn $Script:DataGrid 'IP'          'IPAddress'   120
Add-DgColumn $Script:DataGrid 'MAC'         'MAC'         140
Add-DgColumn $Script:DataGrid 'Nombre'      'Nombre'      160 $true
Add-DgColumn $Script:DataGrid 'Ámbito'      'ScopeId'     115
Add-DgColumn $Script:DataGrid 'Estado'      'Estado'      90
Add-DgColumn $Script:DataGrid 'Descripción' 'Descripcion' 180 $true
Add-DgColumn $Script:DataGrid 'Expira'      'Expira'      130

# Montar panel derecho
$rightPanel.Children.Add($searchCard)   | Out-Null
$rightPanel.Children.Add($filterBar)    | Out-Null
$rightPanel.Children.Add($gridWrapper)  | Out-Null
$rightPanel.Children.Add($statusBar)    | Out-Null

# Montar content
$contentGrid.Children.Add($leftPanel)    | Out-Null
$contentGrid.Children.Add($rightPanel)   | Out-Null
$contentBorder.Child = $contentGrid

# Montar todo
$rootGrid.Children.Add($barraFondo)    | Out-Null
$rootGrid.Children.Add($contentBorder) | Out-Null
$Script:Window.Content = $rootGrid

#==============================================================================
# MÓDULO: Datos en memoria y estado
#==============================================================================
$Script:TodaLaData         = [System.Collections.Generic.List[object]]::new()
$Script:DatosFiltrados     = [System.Collections.Generic.List[object]]::new()
$Script:Cache              = [System.Collections.Generic.List[object]]::new()  # caché global
$Script:CacheCompleto      = $false    # true cuando la carga en background ha terminado
$Script:AmbitosOriginales  = [System.Collections.Generic.List[object]]::new()  # para filtrar lista
$Script:AmbitoSeleccionado = $null
$Script:ModoActual         = 'Busqueda'   # 'Busqueda' | 'Ambito'
$Script:ServidorActivo     = $Script:Config.ServidorPorDefecto
$Script:FiltroBusqueda     = ''

#==============================================================================
# MÓDULO: Helpers de estado UI
#==============================================================================
function Set-Estado {
    param(
        [string]$Texto,
        [string]$Tipo = 'Info'   # Info | Ok | Error | Warning | Loading
    )
    $colores = @{
        'Info'    = $Script:Tema.EstadoInfo
        'Ok'      = $Script:Tema.EstadoOk
        'Error'   = $Script:Tema.EstadoError
        'Warning' = $Script:Tema.EstadoWarning
        'Loading' = $Script:Tema.EstadoCargando
    }
    $color = $colores[$Tipo]
    $Script:LblEstado.Text       = $Texto
    $Script:LblEstado.Foreground = New-WpfBrush $color
    $Script:DotEstado.Fill       = New-WpfBrush $color
    $Script:DotConexion.Fill     = New-WpfBrush $color
}

function Show-Spinner {
    param([string]$Mensaje = 'Consultando servidor DHCP...')
    $Script:LblOverlay.Text         = $Mensaje
    $Script:PanelOverlay.Visibility = 'Visible'
    $Script:SpinnerTimer.Start()
    $Script:BtnBuscar.IsEnabled          = $false
    $Script:BtnRecargarAmbitos.IsEnabled = $false
    $Script:ListaAmbitos.IsHitTestVisible = $false   # bloquear clicks sin cambiar apariencia
}

function Hide-Spinner {
    $Script:PanelOverlay.Visibility = 'Collapsed'
    $Script:SpinnerTimer.Stop()
    $Script:BtnBuscar.IsEnabled          = $true
    $Script:BtnRecargarAmbitos.IsEnabled = $true
    $Script:ListaAmbitos.IsHitTestVisible = $true
}

function Update-Grid {
    param([System.Collections.Generic.List[object]]$Datos)

    $filtroRapido = $Script:TxtFiltroRapido.Text.Trim()
    $placeholder  = $Script:TxtFiltroRapido.Tag
    $tipoFiltro = if ($Script:DdTipoBusqueda.State.SelectedItem) {
        $Script:DdTipoBusqueda.State.SelectedItem.ToString().Trim()
    } else { 'Todo' }

    # Filtrar con Where-Object nativo de PowerShell
    $datosFinal = $Datos | Where-Object {
        ($tipoFiltro -eq 'Todo' -or $_.Tipo -eq $tipoFiltro)
    }

    # Filtro rápido visual
    if (-not [string]::IsNullOrWhiteSpace($filtroRapido) -and $filtroRapido -ne $placeholder) {
        $fr = $filtroRapido
        $datosFinal = $datosFinal | Where-Object {
            $_.IPAddress   -like "*$fr*" -or
            $_.MAC         -like "*$fr*" -or
            $_.Nombre      -like "*$fr*" -or
            $_.Descripcion -like "*$fr*" -or
            $_.ScopeId     -like "*$fr*"
        }
    }

    # Convertir a lista para el DataGrid
    $lista = @($datosFinal)
    $Script:DataGrid.ItemsSource = $lista
    $n = $lista.Count
    $Script:LblConteo.Text = if ($n -gt 0) { "$n resultado$(if($n -ne 1){'s'})" } else { '' }
}

#==============================================================================
# MÓDULO: Recarga de ámbito activo
#==============================================================================
function Reload-CurrentScope {
    if ($null -eq $Script:AmbitoSeleccionado) { return }
    Set-Estado "Recargando ámbito $Script:AmbitoSeleccionado..." 'Loading'
    Show-Spinner "Recargando ámbito $Script:AmbitoSeleccionado..."
    $Script:TodaLaData.Clear()

    Invoke-AsyncDhcp -Trabajo {
        param($Servidor, $ScopeId)
        Get-DhcpScopeDetail -Servidor $Servidor -ScopeId $ScopeId
    } -Argumentos @{ Servidor = $Script:ServidorActivo; ScopeId = $Script:AmbitoSeleccionado } `
    -OnComplete {
        param($Output)
        Hide-Spinner
        $Script:TodaLaData.Clear()
        foreach ($r in $Output) { $Script:TodaLaData.Add($r) | Out-Null }
        $res    = @($Output | Where-Object { $_.Tipo -eq 'Reserva' }).Count
        $leases = @($Output | Where-Object { $_.Tipo -eq 'Lease' }).Count
        $Script:BadgeScopeInfo.Text = "$res reservas · $leases leases"
        Update-Grid $Script:TodaLaData
        Set-Estado "Ámbito $($Script:AmbitoSeleccionado): $($Output.Count) entradas" 'Ok'
    } -OnError {
        param($Err)
        Hide-Spinner
        Set-Estado "Error al recargar ámbito: $Err" 'Error'
    }
}

#==============================================================================
# MÓDULO: Formularios de gestión (reservas y ámbitos)
#==============================================================================
function Show-ReservationForm {
    param(
        [string]$ScopeId,
        [string]$Servidor,
        [string]$Mode = 'New',
        [string]$IP = '',
        [string]$MAC = '',
        [string]$Nombre = '',
        [string]$Descripcion = ''
    )

    $dialog = New-Object System.Windows.Window
    $titleText = if ($Mode -eq 'Edit') { 'Editar reserva' } else { 'Nueva reserva' }
    $dialog.Title  = $titleText
    $dialog.Width           = 460
    $dialog.SizeToContent   = 'Height'
    $dialog.WindowStartupLocation = 'CenterOwner'
    $dialog.Owner      = $Script:Window
    $dialog.Background = New-WpfBrush $Script:Tema.FondoApp
    $dialog.ResizeMode = 'NoResize'
    $dialog.FontFamily = New-Object System.Windows.Media.FontFamily('Segoe UI')

    $mainPanel = New-Object System.Windows.Controls.StackPanel
    $mainPanel.Margin = New-WpfThickness 24 20 24 24

    $dlgTitle = New-Label $titleText $Script:Tema.TextoPrimario 16 'SemiBold'
    $dlgTitle.Margin = New-WpfThickness 0 0 0 4
    $mainPanel.Children.Add($dlgTitle) | Out-Null

    $dlgSubtitle = New-Label "Ámbito: $ScopeId" $Script:Tema.TextoCian 12
    $dlgSubtitle.Margin = New-WpfThickness 0 0 0 16
    $mainPanel.Children.Add($dlgSubtitle) | Out-Null

    $fieldDefs = @(
        @{ Label = 'Dirección IP *'; Value = $IP; Placeholder = 'Ej: 10.95.4.50' },
        @{ Label = 'MAC Address *'; Value = $MAC; Placeholder = 'Ej: 00-1A-2B-3C-4D-5E' },
        @{ Label = 'Nombre'; Value = $Nombre; Placeholder = 'Nombre del equipo (opcional)' },
        @{ Label = 'Descripción'; Value = $Descripcion; Placeholder = 'Descripción (opcional)' }
    )

    $textBoxes = @()
    foreach ($fd in $fieldDefs) {
        $lbl = New-Label $fd.Label $Script:Tema.TextoSecundario 11
        $lbl.Margin = New-WpfThickness 0 0 0 4
        $mainPanel.Children.Add($lbl) | Out-Null

        $tb = New-StyledTextBox -Placeholder $fd.Placeholder -Height 34 -FontSize 12.5
        if (-not [string]::IsNullOrWhiteSpace($fd.Value)) {
            $tb.Text = $fd.Value
            $tb.Foreground = New-WpfBrush $Script:Tema.TextoPrimario
        }
        $tb.Margin = New-WpfThickness 0 0 0 10
        $mainPanel.Children.Add($tb) | Out-Null
        $textBoxes += $tb
    }

    $btnPanel = New-Object System.Windows.Controls.StackPanel
    $btnPanel.Orientation = 'Horizontal'
    $btnPanel.HorizontalAlignment = 'Right'
    $btnPanel.Margin = New-WpfThickness 0 8 0 0

    $dlgBtnCancelar = New-StyledButton -Text 'Cancelar' `
        -BgColor $Script:Tema.FondoInput -FgColor $Script:Tema.TextoSecundario `
        -HoverColor $Script:Tema.Borde -Width 90 -Height 34
    $dlgBtnGuardar = New-StyledButton -Text 'Guardar' -Width 90 -Height 34
    $dlgBtnGuardar.Margin = New-WpfThickness 8 0 0 0

    $btnPanel.Children.Add($dlgBtnCancelar) | Out-Null
    $btnPanel.Children.Add($dlgBtnGuardar) | Out-Null
    $mainPanel.Children.Add($btnPanel) | Out-Null

    $result = [PSCustomObject]@{ Saved = $false; IP = ''; MAC = ''; Nombre = ''; Descripcion = '' }

    $dlgBtnCancelar.Tag = [PSCustomObject]@{
        Normal = $dlgBtnCancelar.Tag.Normal; Hover = $dlgBtnCancelar.Tag.Hover; Press = $dlgBtnCancelar.Tag.Press
        Dialog = $dialog
    }
    $dlgBtnCancelar.Add_Click({ $this.Tag.Dialog.DialogResult = $false; $this.Tag.Dialog.Close() })

    $dlgBtnGuardar.Tag = [PSCustomObject]@{
        Normal = $dlgBtnGuardar.Tag.Normal; Hover = $dlgBtnGuardar.Tag.Hover; Press = $dlgBtnGuardar.Tag.Press
        Dialog = $dialog; TextBoxes = $textBoxes; Result = $result
    }
    $dlgBtnGuardar.Add_Click({
        $ctx = $this.Tag
        $ipVal   = $ctx.TextBoxes[0].Text.Trim()
        $macVal  = $ctx.TextBoxes[1].Text.Trim()
        $nomVal  = $ctx.TextBoxes[2].Text.Trim()
        $descVal = $ctx.TextBoxes[3].Text.Trim()

        if ([string]::IsNullOrWhiteSpace($ipVal) -or $ipVal -eq $ctx.TextBoxes[0].Tag) {
            [System.Windows.MessageBox]::Show('La dirección IP es obligatoria.', 'Validación', 'OK', 'Warning')
            return
        }
        if ([string]::IsNullOrWhiteSpace($macVal) -or $macVal -eq $ctx.TextBoxes[1].Tag) {
            [System.Windows.MessageBox]::Show('La dirección MAC es obligatoria.', 'Validación', 'OK', 'Warning')
            return
        }
        if ($nomVal -eq $ctx.TextBoxes[2].Tag) { $nomVal = '' }
        if ($descVal -eq $ctx.TextBoxes[3].Tag) { $descVal = '' }

        $ctx.Result.Saved = $true
        $ctx.Result.IP = $ipVal
        $ctx.Result.MAC = $macVal
        $ctx.Result.Nombre = $nomVal
        $ctx.Result.Descripcion = $descVal
        $ctx.Dialog.DialogResult = $true
        $ctx.Dialog.Close()
    })

    $dialog.Content = $mainPanel
    $dialog.ShowDialog() | Out-Null
    return $result
}

function Show-ScopeForm {
    param([string]$Servidor)

    $dialog = New-Object System.Windows.Window
    $dialog.Title  = 'Nuevo ámbito DHCP'
    $dialog.Width           = 460
    $dialog.SizeToContent   = 'Height'
    $dialog.WindowStartupLocation = 'CenterOwner'
    $dialog.Owner      = $Script:Window
    $dialog.Background = New-WpfBrush $Script:Tema.FondoApp
    $dialog.ResizeMode = 'NoResize'
    $dialog.FontFamily = New-Object System.Windows.Media.FontFamily('Segoe UI')

    $mainPanel = New-Object System.Windows.Controls.StackPanel
    $mainPanel.Margin = New-WpfThickness 24 20 24 24

    $dlgTitle = New-Label 'Nuevo ámbito DHCP' $Script:Tema.TextoPrimario 16 'SemiBold'
    $dlgTitle.Margin = New-WpfThickness 0 0 0 4
    $mainPanel.Children.Add($dlgTitle) | Out-Null

    $dlgSubtitle = New-Label "Servidor: $Servidor" $Script:Tema.TextoCian 12
    $dlgSubtitle.Margin = New-WpfThickness 0 0 0 16
    $mainPanel.Children.Add($dlgSubtitle) | Out-Null

    $fieldDefs = @(
        @{ Label = 'Red (ScopeId) *'; Placeholder = 'Ej: 10.95.5.0' },
        @{ Label = 'Máscara de subred *'; Placeholder = 'Ej: 255.255.255.0' },
        @{ Label = 'Nombre del ámbito *'; Placeholder = 'Ej: Muros - Planta 1' },
        @{ Label = 'Rango inicio *'; Placeholder = 'Ej: 10.95.5.10' },
        @{ Label = 'Rango fin *'; Placeholder = 'Ej: 10.95.5.200' }
    )

    $textBoxes = @()
    foreach ($fd in $fieldDefs) {
        $lbl = New-Label $fd.Label $Script:Tema.TextoSecundario 11
        $lbl.Margin = New-WpfThickness 0 0 0 4
        $mainPanel.Children.Add($lbl) | Out-Null

        $tb = New-StyledTextBox -Placeholder $fd.Placeholder -Height 34 -FontSize 12.5
        $tb.Margin = New-WpfThickness 0 0 0 10
        $mainPanel.Children.Add($tb) | Out-Null
        $textBoxes += $tb
    }

    $btnPanel = New-Object System.Windows.Controls.StackPanel
    $btnPanel.Orientation = 'Horizontal'
    $btnPanel.HorizontalAlignment = 'Right'
    $btnPanel.Margin = New-WpfThickness 0 8 0 0

    $dlgBtnCancelar = New-StyledButton -Text 'Cancelar' `
        -BgColor $Script:Tema.FondoInput -FgColor $Script:Tema.TextoSecundario `
        -HoverColor $Script:Tema.Borde -Width 90 -Height 34
    $dlgBtnCrear = New-StyledButton -Text 'Crear' `
        -BgColor $Script:Tema.Verde -HoverColor $Script:Tema.VerdeHover `
        -Width 90 -Height 34
    $dlgBtnCrear.Margin = New-WpfThickness 8 0 0 0

    $btnPanel.Children.Add($dlgBtnCancelar) | Out-Null
    $btnPanel.Children.Add($dlgBtnCrear) | Out-Null
    $mainPanel.Children.Add($btnPanel) | Out-Null

    $result = [PSCustomObject]@{ Created = $false; ScopeId = ''; Mascara = ''; Nombre = ''; RangoInicio = ''; RangoFin = '' }

    $dlgBtnCancelar.Tag = [PSCustomObject]@{
        Normal = $dlgBtnCancelar.Tag.Normal; Hover = $dlgBtnCancelar.Tag.Hover; Press = $dlgBtnCancelar.Tag.Press
        Dialog = $dialog
    }
    $dlgBtnCancelar.Add_Click({ $this.Tag.Dialog.DialogResult = $false; $this.Tag.Dialog.Close() })

    $dlgBtnCrear.Tag = [PSCustomObject]@{
        Normal = $dlgBtnCrear.Tag.Normal; Hover = $dlgBtnCrear.Tag.Hover; Press = $dlgBtnCrear.Tag.Press
        Dialog = $dialog; TextBoxes = $textBoxes; Result = $result
    }
    $dlgBtnCrear.Add_Click({
        $ctx = $this.Tag
        $valid = $true
        for ($i = 0; $i -lt 5; $i++) {
            $v = $ctx.TextBoxes[$i].Text.Trim()
            if ([string]::IsNullOrWhiteSpace($v) -or $v -eq $ctx.TextBoxes[$i].Tag) { $valid = $false; break }
        }
        if (-not $valid) {
            [System.Windows.MessageBox]::Show('Todos los campos son obligatorios.', 'Validación', 'OK', 'Warning')
            return
        }

        $ctx.Result.Created    = $true
        $ctx.Result.ScopeId    = $ctx.TextBoxes[0].Text.Trim()
        $ctx.Result.Mascara    = $ctx.TextBoxes[1].Text.Trim()
        $ctx.Result.Nombre     = $ctx.TextBoxes[2].Text.Trim()
        $ctx.Result.RangoInicio = $ctx.TextBoxes[3].Text.Trim()
        $ctx.Result.RangoFin   = $ctx.TextBoxes[4].Text.Trim()
        $ctx.Dialog.DialogResult = $true
        $ctx.Dialog.Close()
    })

    $dialog.Content = $mainPanel
    $dialog.ShowDialog() | Out-Null
    return $result
}

#==============================================================================
# MÓDULO: Lógica de eventos
#==============================================================================

# ── Recargar ámbitos ──────────────────────────────────────────────────────────
$Script:BtnRecargarAmbitos.Add_Click({
    $Script:ServidorActivo = $Script:CmbServidor.SelectedItem
    Set-Estado "Cargando ámbitos desde $Script:ServidorActivo..." 'Loading'
    Show-Spinner 'Cargando ámbitos DHCP...'
    $Script:ListaAmbitos.Items.Clear()
    $Script:BadgeAmbitos.Text = ' 0'

    Invoke-AsyncDhcp -Trabajo {
        param($Servidor)
        Get-DhcpScopes -Servidor $Servidor
    } -Argumentos @{ Servidor = $Script:ServidorActivo } `
    -OnComplete {
        param($Output)
        Hide-Spinner
        $Script:ListaAmbitos.Items.Clear()
        $count = 0
        foreach ($scope in $Output) {
            $item = New-Object System.Windows.Controls.StackPanel
            $item.Tag = $scope.ScopeId.ToString()

            $linea1 = New-Object System.Windows.Controls.TextBlock
            $linea1.Text       = $scope.ScopeId.ToString()
            $linea1.Foreground = New-WpfBrush $Script:Tema.TextoAccent
            Set-WpfFont $linea1 12 'SemiBold'

            $linea2 = New-Object System.Windows.Controls.TextBlock
            $linea2.Text       = if ($scope.Name) { $scope.Name } else { '(sin nombre)' }
            $linea2.Foreground = New-WpfBrush $Script:Tema.TextoSecundario
            Set-WpfFont $linea2 11

            $estadoBadge = New-Object System.Windows.Controls.TextBlock
            $estadoColor = if ($scope.State -eq 'Active') { $Script:Tema.TextoVerde } else { $Script:Tema.TextoAmbar }
            $estadoBadge.Text       = $scope.State
            $estadoBadge.Foreground = New-WpfBrush $estadoColor
            Set-WpfFont $estadoBadge 10

            $item.Children.Add($linea1)      | Out-Null
            $item.Children.Add($linea2)      | Out-Null
            $item.Children.Add($estadoBadge) | Out-Null

            $Script:ListaAmbitos.Items.Add($item) | Out-Null
            $count++
        }
        # Guardar copia de items para filtrado posterior
        $Script:AmbitosOriginales.Clear()
        foreach ($it in $Script:ListaAmbitos.Items) { $Script:AmbitosOriginales.Add($it) | Out-Null }

        $Script:BadgeAmbitos.Text = " $count"
        Set-Estado "Cargados $count ámbitos · cargando datos en background..." 'Loading'
        $Script:DotConexion.Fill = New-WpfBrush $Script:Tema.Verde

        # Lanzar carga completa en background (sin bloquear UI)
        $Script:CacheCompleto = $false
        $Script:Cache.Clear()
        $scopeIds = @($Output | ForEach-Object { $_.ScopeId.ToString() })

        Invoke-AsyncDhcp -Trabajo {
            param($Servidor, $ScopeIds)
            Load-DhcpFullCache -Servidor $Servidor -ScopeIds $ScopeIds
        } -Argumentos @{ Servidor = $Script:ServidorActivo; ScopeIds = $scopeIds } `
        -OnComplete {
            param($Output)
            $Script:Cache.Clear()
            foreach ($r in $Output) { $Script:Cache.Add($r) | Out-Null }
            $Script:CacheCompleto = $true
            $n = $Script:Cache.Count
            Set-Estado "Caché lista: $n entradas de $($Script:ListaAmbitos.Items.Count) ámbitos" 'Ok'
        } -OnError {
            param($Err)
            $Script:CacheCompleto = $false
            Set-Estado "Caché parcial (error en background): $Err" 'Warning'
        }
    } -OnError {
        param($Err)
        Hide-Spinner
        Set-Estado "Error al cargar ámbitos: $Err" 'Error'
    }
})

# ── Selección de ámbito ───────────────────────────────────────────────────────
$Script:ListaAmbitos.Add_SelectionChanged({
    $selected = $Script:ListaAmbitos.SelectedItem
    if ($null -eq $selected) { return }

    $Script:AmbitoSeleccionado = $selected.Tag
    $Script:ServidorActivo     = $Script:CmbServidor.SelectedItem
    $Script:ModoActual = 'Ambito'
    $Script:BtnNuevaReserva.Visibility = 'Visible'

    $Script:LblContexto.Text = "Ámbito: $Script:AmbitoSeleccionado"
    $Script:LblContexto.Foreground = New-WpfBrush $Script:Tema.TextoCian
    $Script:BadgeScopeInfo.Text = 'Cargando reservas y leases...'

    Set-Estado "Cargando ámbito $Script:AmbitoSeleccionado..." 'Loading'
    Show-Spinner "Cargando ámbito $Script:AmbitoSeleccionado..."
    $Script:TodaLaData.Clear()

    Invoke-AsyncDhcp -Trabajo {
        param($Servidor, $ScopeId)
        Get-DhcpScopeDetail -Servidor $Servidor -ScopeId $ScopeId
    } -Argumentos @{ Servidor = $Script:ServidorActivo; ScopeId = $Script:AmbitoSeleccionado } `
    -OnComplete {
        param($Output)
        Hide-Spinner
        $Script:TodaLaData.Clear()
        foreach ($r in $Output) { $Script:TodaLaData.Add($r) | Out-Null }

        $res    = ($Output | Where-Object { $_.Tipo -eq 'Reserva' }).Count
        $leases = ($Output | Where-Object { $_.Tipo -eq 'Lease' }).Count
        $Script:BadgeScopeInfo.Text = "$res reservas · $leases leases"
        Update-Grid $Script:TodaLaData
        Set-Estado "Ámbito $($Script:AmbitoSeleccionado): $($Output.Count) entradas" 'Ok'
    } -OnError {
        param($Err)
        Hide-Spinner
        Set-Estado "Error al cargar ámbito: $Err" 'Error'
        $Script:BadgeScopeInfo.Text = ''
    }
})

# ── Búsqueda global ───────────────────────────────────────────────────────────
$Script:BtnBuscar.Add_Click({
    $Script:FiltroBusqueda = $Script:TxtBuscar.Text.Trim()
    $placeholder           = $Script:TxtBuscar.Tag

    if ([string]::IsNullOrWhiteSpace($Script:FiltroBusqueda) -or $Script:FiltroBusqueda -eq $placeholder) {
        Set-Estado 'Introduce un término de búsqueda' 'Warning'
        $Script:TxtBuscar.Focus() | Out-Null
        return
    }

    $Script:ServidorActivo = $Script:CmbServidor.SelectedItem
    $Script:ModoActual = 'Busqueda'
    $Script:BtnNuevaReserva.Visibility = 'Collapsed'
    $Script:ListaAmbitos.SelectedIndex = -1
    $Script:LblContexto.Text = "Búsqueda: '$Script:FiltroBusqueda'"
    $Script:LblContexto.Foreground = New-WpfBrush $Script:Tema.Accent

    Set-Estado "Buscando '$Script:FiltroBusqueda' en $Script:ServidorActivo..." 'Loading'
    Show-Spinner "Buscando '$Script:FiltroBusqueda' en todos los ámbitos..."
    $Script:TodaLaData.Clear()
    $Script:DataGrid.ItemsSource = $null

    # Si la caché está disponible, buscar en local (instantáneo)
    if ($Script:CacheCompleto -and $Script:Cache.Count -gt 0) {
        $fr = $Script:FiltroBusqueda
        $Script:TodaLaData.Clear()
        $Script:Cache | Where-Object {
            $_.IPAddress   -like "*$fr*" -or
            $_.MAC         -like "*$fr*" -or
            $_.Nombre      -like "*$fr*" -or
            $_.Descripcion -like "*$fr*" -or
            $_.ScopeId     -like "*$fr*"
        } | ForEach-Object { $Script:TodaLaData.Add($_) | Out-Null }

        $Script:DataGrid.ItemsSource = $null
        Update-Grid $Script:TodaLaData
        $n = $Script:TodaLaData.Count
        Hide-Spinner
        if ($n -eq 0) {
            Set-Estado "Sin resultados para '$fr' (búsqueda local)" 'Warning'
        } else {
            Set-Estado "Encontradas $n entradas para '$fr' (búsqueda local)" 'Ok'
        }
        return
    }

    # Sin caché: consulta directa al servidor
    Invoke-AsyncDhcp -Trabajo {
        param($Servidor, $Filtro)
        Search-DhcpAll -Servidor $Servidor -Filtro $Filtro
    } -Argumentos @{ Servidor = $Script:ServidorActivo; Filtro = $Script:FiltroBusqueda } `
    -OnComplete {
        param($Output)
        Hide-Spinner
        $Script:TodaLaData.Clear()
        foreach ($r in $Output) { $Script:TodaLaData.Add($r) | Out-Null }
        Update-Grid $Script:TodaLaData
        $n = $Script:TodaLaData.Count
        if ($n -eq 0) {
            Set-Estado "Sin resultados para '$Script:FiltroBusqueda'" 'Warning'
        } else {
            Set-Estado "Encontradas $n entradas para '$Script:FiltroBusqueda'" 'Ok'
        }
    } -OnError {
        param($Err)
        Hide-Spinner
        Set-Estado "Error en búsqueda: $Err" 'Error'
    }
})

# ── Tecla Enter en búsqueda ───────────────────────────────────────────────────
$Script:TxtBuscar.Add_KeyDown({
    if ($_.Key -eq 'Return') { $Script:BtnBuscar.RaiseEvent(
        (New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
})

# ── Filtro rápido en tiempo real ──────────────────────────────────────────────
$Script:TxtFiltroRapido.Add_TextChanged({
    if ($Script:TodaLaData.Count -gt 0) {
        Update-Grid $Script:TodaLaData
    }
})

# ── Selector de tipo búsqueda ─────────────────────────────────────────────────
$Script:DdTipoBusqueda.State.Add_SelectionChanged({
    if ($Script:TodaLaData.Count -gt 0) {
        Update-Grid $Script:TodaLaData
    }
})

# ── Filtro de ámbitos en tiempo real ─────────────────────────────────────────
$Script:TxtFiltroAmbitos.Add_TextChanged({
    $texto = $this.Text.Trim()
    $ph    = $this.Tag
    if ($Script:AmbitosOriginales.Count -eq 0) { return }

    $Script:ListaAmbitos.Items.Clear()
    $filtrados = if ([string]::IsNullOrWhiteSpace($texto) -or $texto -eq $ph) {
        $Script:AmbitosOriginales
    } else {
        $Script:AmbitosOriginales | Where-Object {
            # Tag contiene el ScopeId, el StackPanel tiene TextBlock hijos con nombre y estado
            $scopeId = $_.Tag
            $nombre  = if ($_.Children.Count -ge 2) { $_.Children[1].Text } else { '' }
            $scopeId -like "*$texto*" -or $nombre -like "*$texto*"
        }
    }
    foreach ($it in $filtrados) { $Script:ListaAmbitos.Items.Add($it) | Out-Null }
    $Script:BadgeAmbitos.Text = " $($Script:ListaAmbitos.Items.Count)"
})

# ── Botón limpiar ─────────────────────────────────────────────────────────────
$Script:BtnLimpiar.Add_Click({
    $Script:TxtBuscar.Text = $Script:TxtBuscar.Tag
    $Script:TxtBuscar.Foreground = New-WpfBrush $Script:Tema.TextoTerciario
    $Script:TxtFiltroRapido.Text = $Script:TxtFiltroRapido.Tag
    $Script:TxtFiltroRapido.Foreground = New-WpfBrush $Script:Tema.TextoTerciario
    $Script:ListaAmbitos.SelectedIndex = -1
    $Script:DataGrid.ItemsSource = $null
    $Script:TodaLaData.Clear()
    $Script:LblContexto.Text = 'Todos los ámbitos'
    $Script:LblContexto.Foreground = New-WpfBrush $Script:Tema.TextoSecundario
    $Script:LblConteo.Text = ''
    $Script:BadgeScopeInfo.Text = ''
    $Script:ModoActual = 'Busqueda'
    $Script:BtnNuevaReserva.Visibility = 'Collapsed'
    Set-Estado 'Listo' 'Info'
})

# ── Cambio de servidor ────────────────────────────────────────────────────────
$Script:DdServidor.State.Add_SelectionChanged({
    $Script:BtnLimpiar.RaiseEvent(
        (New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    $Script:ListaAmbitos.Items.Clear()
    $Script:BadgeAmbitos.Text = ' 0'
    $Script:DotConexion.Fill = New-WpfBrush $Script:Tema.TextoTerciario
    Set-Estado "Servidor cambiado a $($Script:DdServidor.State.SelectedItem)" 'Info'
})

# ── Nueva reserva ────────────────────────────────────────────────────────────
$Script:BtnNuevaReserva.Add_Click({
    if ($null -eq $Script:AmbitoSeleccionado) {
        Set-Estado 'Selecciona un ámbito primero' 'Warning'
        return
    }

    $result = Show-ReservationForm -ScopeId $Script:AmbitoSeleccionado -Servidor $Script:ServidorActivo -Mode 'New'

    if ($result.Saved) {
        $Script:NewReservaResult = $result
        Set-Estado "Creando reserva $($result.IP)..." 'Loading'
        Show-Spinner 'Creando reserva en el servidor...'

        Invoke-AsyncDhcp -Trabajo {
            param($Servidor, $ScopeId, $IP, $MAC, $Nombre, $Descripcion)
            New-DhcpReservation -Servidor $Servidor -ScopeId $ScopeId -IP $IP -MAC $MAC -Nombre $Nombre -Descripcion $Descripcion
        } -Argumentos @{
            Servidor = $Script:ServidorActivo; ScopeId = $Script:AmbitoSeleccionado
            IP = $result.IP; MAC = $result.MAC; Nombre = $result.Nombre; Descripcion = $result.Descripcion
        } -OnComplete {
            param($Output)
            Hide-Spinner
            Set-Estado "Reserva creada: $($Script:NewReservaResult.IP)" 'Ok'
            Reload-CurrentScope
        } -OnError {
            param($Err)
            Hide-Spinner
            Set-Estado "Error al crear reserva: $Err" 'Error'
        }
    }
})

# ── Nuevo ámbito ─────────────────────────────────────────────────────────────
$Script:BtnNuevoAmbito.Add_Click({
    $Script:ServidorActivo = $Script:CmbServidor.SelectedItem
    $result = Show-ScopeForm -Servidor $Script:ServidorActivo

    if ($result.Created) {
        Set-Estado "Creando ámbito $($result.ScopeId)..." 'Loading'
        Show-Spinner 'Creando ámbito en el servidor...'

        Invoke-AsyncDhcp -Trabajo {
            param($Servidor, $ScopeId, $Mascara, $Nombre, $RangoInicio, $RangoFin)
            New-DhcpScopeCustom -Servidor $Servidor -ScopeId $ScopeId -Mascara $Mascara `
                -Nombre $Nombre -RangoInicio $RangoInicio -RangoFin $RangoFin
        } -Argumentos @{
            Servidor = $Script:ServidorActivo; ScopeId = $result.ScopeId
            Mascara = $result.Mascara; Nombre = $result.Nombre
            RangoInicio = $result.RangoInicio; RangoFin = $result.RangoFin
        } -OnComplete {
            param($Output)
            Hide-Spinner
            Set-Estado 'Ámbito creado correctamente. Recargando...' 'Ok'
            $Script:BtnRecargarAmbitos.RaiseEvent(
                (New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
        } -OnError {
            param($Err)
            Hide-Spinner
            Set-Estado "Error al crear ámbito: $Err" 'Error'
        }
    }
})

# ── Doble clic: editar reserva o copiar lease ─────────────────────────────────
$Script:DataGrid.Add_MouseDoubleClick({
    try {
        $item = $Script:DataGrid.SelectedItem
        if ($null -eq $item) { return }

        if ($item.Tipo -eq 'Reserva') {
            $Script:EditOriginalIP = $item.IPAddress
            $Script:EditScopeId    = $item.ScopeId

            $result = Show-ReservationForm -ScopeId $item.ScopeId -Servidor $Script:ServidorActivo `
                -Mode 'Edit' -IP $item.IPAddress -MAC $item.MAC -Nombre $item.Nombre -Descripcion $item.Descripcion

            if ($result.Saved) {
                $Script:EditResult = $result
                Set-Estado "Editando reserva $($result.IP)..." 'Loading'
                Show-Spinner 'Editando reserva en el servidor...'

                Invoke-AsyncDhcp -Trabajo {
                    param($Servidor, $ScopeId, $OldIP, $IP, $MAC, $Nombre, $Descripcion)
                    Edit-DhcpReservation -Servidor $Servidor -ScopeId $ScopeId -OldIP $OldIP `
                        -IP $IP -MAC $MAC -Nombre $Nombre -Descripcion $Descripcion
                } -Argumentos @{
                    Servidor = $Script:ServidorActivo; ScopeId = $Script:EditScopeId
                    OldIP = $Script:EditOriginalIP; IP = $result.IP; MAC = $result.MAC
                    Nombre = $result.Nombre; Descripcion = $result.Descripcion
                } -OnComplete {
                    param($Output)
                    Hide-Spinner
                    Set-Estado "Reserva editada: $($Script:EditResult.IP)" 'Ok'
                    if ($Script:ModoActual -eq 'Ambito') { Reload-CurrentScope }
                } -OnError {
                    param($Err)
                    Hide-Spinner
                    Set-Estado "Error al editar reserva: $Err" 'Error'
                }
            }
        } else {
            $texto = "IP: $($item.IPAddress) | MAC: $($item.MAC) | Nombre: $($item.Nombre) | Ámbito: $($item.ScopeId)"
            [System.Windows.Forms.Clipboard]::SetText($texto)
            Set-Estado "Copiado al portapapeles: $($item.IPAddress)" 'Ok'
        }
    } catch {
        try { Hide-Spinner } catch {}
        Set-Estado "Error: $($_.Exception.Message)" 'Error'
    }
})

# ── Copiar fila con Ctrl+C ───────────────────────────────────────────────────
$Script:DataGrid.Add_KeyDown({
    if ($_.Key -eq 'C' -and [System.Windows.Input.Keyboard]::IsKeyDown('LeftCtrl')) {
        try {
            $item = $Script:DataGrid.SelectedItem
            if ($null -eq $item) { return }
            $texto = "$($item.IPAddress)`t$($item.MAC)`t$($item.Nombre)`t$($item.ScopeId)`t$($item.Estado)"
            [System.Windows.Forms.Clipboard]::SetText($texto)
            Set-Estado "Copiado: $($item.IPAddress)" 'Ok'
        } catch {}
    }
})

#==============================================================================
# MÓDULO: Arranque
#==============================================================================
$Script:Window.Add_Loaded({
    Set-Estado "Conectando a $($Script:Config.ServidorPorDefecto)..." 'Loading'
    # Auto-cargar ámbitos al arrancar
    $Script:BtnRecargarAmbitos.RaiseEvent(
        (New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
})

$Script:Window.Add_Closed({
    if ($Script:SpinnerTimer.IsEnabled) { $Script:SpinnerTimer.Stop() }
})

# ── Arrancar UI ───────────────────────────────────────────────────────────────
Write-Host "[$($Script:Config.TituloApp) v$($Script:Config.Version)] Iniciando..."
Write-Host "[+] Servidor por defecto : $($Script:Config.ServidorPorDefecto)"
Write-Host "[+] Motor                : WPF + Runspaces"

$Script:Window.ShowDialog() | Out-Null