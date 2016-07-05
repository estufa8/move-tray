Imports System.Security.Permissions
Imports System.ComponentModel.Design
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Forms.Design
Imports System.Windows

'''<summary>
'''  Interaction logic for MyControl.xaml
'''</summary>
Partial Public Class MoveTray
    Inherits System.Windows.Controls.UserControl
    Public dte As EnvDTE80.DTE2
    Public design As IDesignerHost

    Public WriteOnly Property Comentaris As String
        Set(value As String)
            Me.txComentaris.Text = value
        End Set
    End Property

    Private Sub btnBuscarControlsOrfes_Click(sender As Object, e As Windows.RoutedEventArgs) Handles btnBuscarControlsOrfes.Click
        'busca controles huérfanos. Que están declarados (primer regex) pero que no están en líneas que previamente contengan Add, AddRange, .Datasource=, etc + nombre control

        If design IsNot Nothing AndAlso dte IsNot Nothing Then
            Dim sarch As String = dte.ActiveDocument.FullName.Replace(".vb", ".Designer.vb")
            If Not IO.File.Exists(sarch) Then
                MsgBox("Hem de tenir activat un formulari en mode disseny per executar això.")
                Return
            End If
            lbLlista.Visibility = Windows.Visibility.Hidden
            lbLlista.ItemsSource = Nothing
            txComentaris.Visibility = Windows.Visibility.Visible
            txComentaris.Text = "Un moment si us plau, buscant controls orfes..."
            txComentaris.UpdateLayout()
            Dim sCod As String = IO.File.ReadAllText(sarch, System.Text.Encoding.Default)
            Dim lstTodos As New List(Of String)
            For Each m As Match In Regex.Matches(sCod, "(?<=Friend WithEvents )\S+", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
                lstTodos.Add(m.Value)
            Next
            Dim spat As String = "(?<=\.add\(|\.addrange\(|\.DataSource = |\.View = |Images = |Repository = )(.+)("
            Dim spatTodos As String = "\b" & String.Join("\b|\b", lstTodos.ToArray) & "\b"
            spat &= spatTodos & ")"
            Dim lstConAdd As New List(Of String)
            For Each m As Match In Regex.Matches(sCod, spat, RegexOptions.Multiline Or RegexOptions.IgnoreCase)
                If m.Groups.Count < 3 Then Continue For
                'pillamos la línea entera. En Expresso pilla cada control de la línea, aquí no. Match ya es true si coincide un control dentro del OR (|)
                'por ejemplo una línea con AddRange(tabGeneral, tabTarifas) devuelve true con tabGeneral, no por cada tab
                Dim slin As String = m.Groups(0).Value
                'una línea puede contener varios controles, por tanto volvemos a filtrar
                For Each m1 As Match In Regex.Matches(slin, spatTodos, RegexOptions.Singleline Or RegexOptions.IgnoreCase Or RegexOptions.Compiled)
                    Dim s As String = m1.Value
                    If Not lstConAdd.Contains(s) Then lstConAdd.Add(s)
                Next
            Next
            Dim lstNoEstan As New List(Of String)
            For Each control As String In lstTodos
                If Not lstConAdd.Contains(control, StringComparer.CurrentCultureIgnoreCase) Then
                    lstNoEstan.Add(control)
                End If
            Next
            If lstNoEstan.Count > 0 Then
                lstNoEstan.Sort()
                txComentaris.Text = String.Join(vbCrLf, lstNoEstan.ToArray)
            End If
        Else
            MsgBox("Hem de tenir activat un formulari en mode disseny per executar això.")
        End If
    End Sub

    Private Sub btnBuscarCadenesControlsPerResx_Click(sender As Object, e As Windows.RoutedEventArgs) Handles btnBuscarCadenesControlsPerResx.Click
        If design IsNot Nothing AndAlso dte IsNot Nothing Then
            Dim sarch As String = dte.ActiveDocument.FullName.Replace(".vb", ".Designer.vb")
            If Not IO.File.Exists(sarch) Then
                MsgBox("Hem de tenir activat un formulari en mode disseny per executar això.")
                Return
            End If
            lbLlista.Visibility = Windows.Visibility.Hidden
            lbLlista.ItemsSource = Nothing
            txComentaris.Visibility = Windows.Visibility.Visible
            txComentaris.Text = "Un moment si us plau, buscant textes (Text|Caption|ToolTip)..."
            txComentaris.UpdateLayout()
            Dim sCod As String = IO.File.ReadAllText(sarch, System.Text.Encoding.Default)
            Dim resx As String = ""
            Dim scont As String = ""
            Dim dict As New Dictionary(Of String, String())
            Dim sExclosos As String = "^Splitcontainer|^Standalonebar|Panel\d+|Barbuttonitem\d"
            For Each m As Match In Regex.Matches(sCod, "(\S+)(\.Text = |\.Caption = |\.ToolTip = )(""[^""]*?"")", RegexOptions.Multiline)
                'tenim 4 grups: 0 amb tot, (1)nom control, (2).caption|.text|etc, (3)"texte"
                Dim lst As New List(Of String) 'guardem propietat resx i texte
                Dim sNomCont As String = m.Groups(1).Value 'Label1
                sNomCont &= m.Groups(2).Value 'nom control  + caption = 
                Dim sText As String = m.Groups(3).Value.Trim("""") 'Dirección
                If sText.Trim.Length = 0 Then Continue For
                Dim sNomResx As String = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(sText.Normalize(Text.NormalizationForm.FormD)) 'a CamelCase y quitando acentos, ñ, etc.
                sNomResx = Regex.Replace(sNomResx, "[^a-z0-9]", "", RegexOptions.IgnoreCase) 'quita lo que no sean letras o números
                If Regex.Match(sNomResx, sExclosos, RegexOptions.IgnoreCase Or RegexOptions.Compiled).Success Then Continue For
                If Char.IsDigit(sNomResx(0)) Then sNomResx = "n" & sNomResx
                lst.Add(sNomResx)
                lst.Add(sText)
                dict.Add(sNomCont, lst.ToArray)
            Next
            'ordenamos por nombre de propiedad, es decir por el texto (pej "NombreComercial")
            dict = (From kv In dict Order By kv.Value(0) Select kv).ToDictionary(Function(k) k.Key, Function(v) v.Value)
            Dim sResult As New StringBuilder
            sResult.AppendLine("Busquem cadenes després de: Caption = |Text = |ToolTip = (avisar si hi ha més casos)")
            sResult.AppendLine("S'afegeixen comentaris a la part de controls per les incidencies aparegudes.")
            sResult.AppendLine()
            sResult.AppendLine("Codi a enganxar al resx:")
            Dim lstConts As New List(Of String)
            Dim sPropAnterior As String = ""
            Dim sValAnterior As String = ""
            Dim indx As Integer = 0
            For Each kv In dict
                Dim bSaltaProp = False
                Dim sComentari As String = ""
                Dim arr() As String = kv.Value '[prop, valor]
                If arr(0).Equals(sPropAnterior) Then
                    If arr(1).Equals(sValAnterior) Then
                        bSaltaProp = True
                        sComentari = " 'Compte: Nom del resource i valor repetits. Fem servir el mateix nom."
                    Else
                        indx += 1 'prop se llama igual +1 (pej NombreComercial1)
                        sComentari = " 'Compte: Nom del resource repetit però diferent valor, afegim " & indx & " al nom."
                    End If
                Else 'como vienen ordenados, aquí ya hemos cambiado de prop
                    indx = 0 'reiniciamos indx de props iguales
                End If
                If indx > 0 Then arr(0) &= indx 'prop igual con distinto valor
                If kv.Key.StartsWith("Me.") Then
                    lstConts.Add(kv.Key & "My.Resources.XXXX." & arr(0) & sComentari)
                Else
                    lstConts.Add("'control sense ""Me."", possiblement amb GenerateMember=False" & vbCrLf & kv.Key & "My.Resources.XXXX." & arr(0) & sComentari)
                End If
                If bSaltaProp Then Continue For 'ja hem posat una prop amb tot igual
                sResult.AppendLine("  <data name=""" & arr(0) & """ xml:space=""preserve"">")
                sResult.AppendLine("    <value>" & arr(1) & "</value>")
                sResult.AppendLine("  </data>")
                sPropAnterior = arr(0)
                sValAnterior = arr(1)
            Next
            lstConts.Sort()
            sResult.AppendLine()
            sResult.AppendLine("Codi a enganxar al form quan canvia l'idioma:")
            sResult.AppendLine()
            sResult.AppendLine(String.Join(vbCrLf, lstConts))
            sResult.AppendLine()
            sResult.AppendLine("No es fa res més. Copiar al notepad, canviar XXXX pel nom del resx i fer servir al gust.")
            txComentaris.Text = sResult.ToString
        Else
            MsgBox("Hem d'estar veient un Form en mode disseny per executar això.")
        End If
    End Sub

    Private Sub btnBuscarCadenesAArxius_Click(sender As Object, e As Windows.RoutedEventArgs) Handles btnBuscarCadenesAArxius.Click
        If design IsNot Nothing AndAlso dte IsNot Nothing Then
            Dim sarch As String = dte.ActiveDocument.FullName.Replace(".vb", ".Designer.vb")
            If Not IO.File.Exists(sarch) Then
                MsgBox("Hem de tenir activat un formulari en mode disseny per executar això.")
                Return
            End If
            lbLlista.Visibility = Windows.Visibility.Visible
            lbLlista.ItemsSource = Nothing
            txComentaris.Visibility = Windows.Visibility.Hidden
            txComentaris.Text = ""
            Dim sCod As String = IO.File.ReadAllText(sarch, System.Text.Encoding.Default)
            Dim patt As String = "(?<=^\s*?Partial\s+Class\s+)\S+"
            Dim clase As String = Regex.Match(sCod, patt, RegexOptions.Multiline Or RegexOptions.IgnoreCase).Value
            Dim ofd As New FolderBrowserDialog
            ofd.SelectedPath = IO.Path.GetDirectoryName(sarch)
            ofd.Description = "Carpeta on cercar arxiu amb 'Partial Class " & clase & "' i cadenes"
            Dim lstItems As New List(Of ItemCadena)
            If ofd.ShowDialog = DialogResult.OK Then
                patt = "^[^']+Class\s+" & clase & "\b"
                For Each arch In My.Computer.FileSystem.GetFiles(ofd.SelectedPath, FileIO.SearchOption.SearchAllSubDirectories, "*.vb")
                    Dim bPosatArxiu As Boolean = False
                    If arch.ToLower = sarch.ToLower Then Continue For
                    sCod = IO.File.ReadAllText(arch, System.Text.Encoding.Default)
                    If Not Regex.Match(sCod, patt, RegexOptions.IgnoreCase Or RegexOptions.Multiline Or RegexOptions.Compiled).Success Then Continue For
                    For Each m As Match In Regex.Matches(sCod, "(?<=^[^']+)(""[^""]*"")", RegexOptions.Multiline Or RegexOptions.Compiled)
                        If m.Groups.Count > 1 Then
                            For i = 1 To m.Groups.Count - 1
                                Dim sg As String = m.Groups(i).Value.Trim(Chr(34))
                                If sg.Length > 0 Then
                                    'si no empieza por letra o número pasamos de ella, pues hay más casos de basura que de textos válidos
                                    If Not Regex.Match(sg.TrimStart, "^[a-z0-9]", RegexOptions.IgnoreCase Or RegexOptions.Compiled).Success Then Continue For
                                    If Not bPosatArxiu Then
                                        If lstItems.Count > 0 Then
                                            lstItems.Add(New ItemCadena)
                                            lstItems.Add(New ItemCadena)
                                        End If
                                        'el 0 indica que es item de archivo
                                        lstItems.Add(New ItemCadena(arch, 0, "Arxiu: " & Regex.Replace(arch, Regex.Escape(ofd.SelectedPath), "", RegexOptions.IgnoreCase)))
                                        bPosatArxiu = True
                                    End If
                                    Dim ilin As Integer = sCod.Substring(0, m.Groups(i).Index).Split(vbCr).Count - 1
                                    lstItems.Add(New ItemCadena(arch, ilin, sg))
                                End If
                            Next
                        End If
                    Next
                Next
                If lstItems.Count > 0 Then
                    lstItems.Insert(0, New ItemCadena("", 0, "")) 'al revés porque insertamos en indx 0
                    lstItems.Insert(0, New ItemCadena("", 0, "També hi ha un menú contextual."))
                    lstItems.Insert(0, New ItemCadena("", 0, "Clicant a un item s'obre l'arxiu i anem a la línia."))
                Else
                    lstItems.Add(New ItemCadena("", 0, "No s'han trobat cadenes."))
                End If
                lbLlista.ItemsSource = lstItems
            End If
            'txComentaris.Text = "Un moment si us plau, buscant textes (Text|Caption|ToolTip)..."
            'txComentaris.UpdateLayout()
        Else
            MsgBox("Hem de tenir activat un formulari en mode disseny per executar això.")
        End If
    End Sub
    Public Class ItemCadena
        Public Archivo As String
        Public Linea As String
        Public Contenido As String
        Public Sub New()

        End Sub
        Public Sub New(arch As String, indx As Integer, cont As String)
            Me.Archivo = arch
            Me.Linea = indx
            Me.Contenido = cont
        End Sub

        Public ReadOnly Property TextoItem() As String
            Get
                Return Contenido
            End Get
        End Property
    End Class

    Private Sub lbLlista_SelectionChanged(sender As Object, e As Windows.Controls.SelectionChangedEventArgs) Handles lbLlista.SelectionChanged
        If lbLlista.SelectedItem Is Nothing Then Return
        Dim it As ItemCadena = lbLlista.SelectedItem
        If it.Linea = 0 Then Return 'es archivo
        Dim w As EnvDTE.Window = dte.ItemOperations.OpenFile(it.Archivo, EnvDTE.Constants.vsViewKindTextView)
        CType(dte.ActiveDocument.Selection, EnvDTE.TextSelection).GotoLine(it.Linea + 1, True)
    End Sub

    Private Sub mnuCrearResxItemsCadenes_Click(sender As Object, e As Windows.RoutedEventArgs) Handles mnuCrearResxItemsCadenes.Click
        If lbLlista.Items.Count = 0 Then Return
        Dim sb As New StringBuilder
        sb.AppendLine("Copiar com texte i usar al gust.")
        sb.AppendLine("Es posen noms de recurs diferents per poder editar amb l'editor de recursos.")
        sb.AppendLine("No s'ordena de cap forma. L'editor de resx ja permet ordenar.")
        sb.AppendLine("Si a l'editor renombrem un recurs ja utilitzat, es renombra al codi, però no al revés (renombrar-lo al codi no afecta al recurs).")
        sb.AppendLine()
        Dim indx As Integer = 1
        For Each it As ItemCadena In lbLlista.ItemsSource
            If it.Linea = 0 Then Continue For 'es archivo
            sb.AppendLine("  <data name=""Item" & indx.ToString("000") & """ xml:space=""preserve"">")
            sb.AppendLine("    <value>" & it.Contenido.Trim & "</value>")
            sb.AppendLine("  </data>")
            indx += 1
        Next
        lbLlista.Visibility = Windows.Visibility.Hidden
        txComentaris.Visibility = Windows.Visibility.Visible
        txComentaris.Text = sb.ToString
    End Sub

    Private Sub btnMoureTrayADreta_Click(sender As Object, e As Windows.RoutedEventArgs) Handles btnMoureTrayADreta.Click
        If design IsNot Nothing AndAlso dte IsNot Nothing Then
            Dim tray As ComponentTray = design.GetService(GetType(ComponentTray))
            If tray Is Nothing Then Return
            If Not tray.Dock = DockStyle.Right Then tray.Dock = DockStyle.Right
            If tray.Width = 0 Then tray.Width = 150
            tray.PerformLayout()
            'ordenamos  los controles alfabeticamente
            Dim dict As New SortedDictionary(Of String, Control)
            For Each c As Control In tray.Controls
                dict.Add(c.Text, c) 'sembla tonto però s'afegeixen ordenats per nom
            Next
            Dim y As Integer = 17
            For Each kv In dict 'ja els tenim ordenats per nom
                kv.Value.Location = New System.Drawing.Point(17, y)
                y += 39 '39 parece ser el separador por defecto (al reordenar con el menú contextual del tray)
            Next
            tray.PerformLayout()
            Dim frame As Control = tray.Parent 'DesignFrame (type oculto) y contiene el Form, el splitter y el tray
            Dim overlay = frame.Controls(0) 'ya es DockStyle.Full
            If overlay.Width = 0 Then overlay.Width = 500 'una vez me ha pasado de abrir el form en diseño y que el ComponentTray ocupe todo el espacio
            Dim spl As Splitter = frame.Controls(1) 'el splitter
            If Not spl.Dock = DockStyle.Right Then spl.Dock = DockStyle.Right
        Else
            MsgBox("Hem de tenir activat un formulari en mode disseny per executar això.")
        End If

    End Sub

    Dim WithEvents frm1024 As Form

    Private Sub chkDibuja1024_Checked(sender As Object, e As Windows.RoutedEventArgs) Handles chkDibuja1024.Checked
        'If chkDibuja1024.IsChecked Then
        '    timerDibuja1024.Interval = 500
        '    timerDibuja1024.Start()
        'Else
        '    timerDibuja1024.Stop()
        'End If
        If chkDibuja1024.IsChecked Then
            If Not frm1024 Is Nothing Then
                frm1024.Close()
                frm1024.Dispose()
                frm1024 = Nothing
            End If
            frm1024 = New Form
            frm1024.Size = New System.Drawing.Size(1280, 1024)
            frm1024.FormBorderStyle = FormBorderStyle.FixedDialog
            frm1024.Opacity = 0.5
            frm1024.TopMost = True
            frm1024.MaximizeBox = False
            frm1024.Text = "Formulari de 1280x1024"
            frm1024.Show()
        End If
    End Sub

    Private Sub frm1024_FormClosed(sender As Object, e As FormClosedEventArgs) Handles frm1024.FormClosed
        chkDibuja1024.IsChecked = False
    End Sub

    Private Sub btnTabulacions_Click(sender As Object, e As Windows.RoutedEventArgs) Handles btnTabulacions.Click
        If design IsNot Nothing AndAlso dte IsNot Nothing Then
            Dim sarch As String = dte.ActiveDocument.FullName.Replace(".vb", ".Designer.vb")
            If Not IO.File.Exists(sarch) Then
                MsgBox("Hem de tenir activat un formulari en mode disseny per executar això.")
                Return
            End If
            Dim scod As String = IO.File.ReadAllText(sarch, System.Text.Encoding.UTF8)
            Dim msConTabIndex = Regex.Matches(scod, "(?<=^\s+)(\S+)(\.TabIndex = )(\d+)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
            Dim lineasconControlsAdd As String = "" 'ahorramos buscar parents en TODO el Designer.vb
            For Each m As Match In Regex.Matches(scod, "\S+\.Add\S*?\(.+", RegexOptions.IgnoreCase Or RegexOptions.Multiline)
                lineasconControlsAdd &= m.Value.Trim & vbCrLf
            Next
            lineasconControlsAdd = lineasconControlsAdd.Trim
            Dim lst As New List(Of ElementoTab)
            'buscamos controles con TabIndex= y nos quedamos con su nombre y con su padre (padre.Controls.Add(...control))
            For Each m As Match In msConTabIndex
                Dim cont As New ElementoTab
                cont.Nombre = m.Groups(1).Value
                cont.TabIndx = m.Groups(3).Value
                cont.NombrePadre = Regex.Match(lineasconControlsAdd, "(\S+)(?:\.[^.]+\.Add\S*?\(.*?" & cont.Nombre & ")", RegexOptions.Multiline Or RegexOptions.IgnoreCase Or RegexOptions.Compiled).Groups(1).Value
                lst.Add(cont)
            Next

            For Each elem As ElementoTab In lst.ToArray
                'añadimos posibles padres que no tienen TabIndex (por ejemplo una pestaña)
                Dim sPadre As String = elem.NombrePadre
                If sPadre = "Me" Then Continue For
                If Regex.Match(sPadre, "\S+\.Panel\d").Success Then
                    'como Panel1/2 no se agregan al splittercontainer, los agregamos a mano al lst
                    Dim el As ElementoTab = lst.FirstOrDefault(Function(x) x.Nombre = sPadre)
                    If el Is Nothing Then
                        el = New ElementoTab
                        el.Nombre = sPadre
                        el.NombrePadre = sPadre.Substring(0, sPadre.IndexOf(".Panel"))
                        el.TabIndx = -1
                        lst.Add(el)
                    End If
                    Continue For
                End If
                Dim elP As ElementoTab = lst.FirstOrDefault(Function(x) x.Nombre = sPadre)
                If elP Is Nothing Then
                    elP = New ElementoTab
                    elP.Nombre = sPadre
                    elP.NombrePadre = Regex.Match(lineasconControlsAdd, "(\S+)(?:\.[^.]+\.Add\S*?\(.*?" & sPadre & ")", RegexOptions.Multiline Or RegexOptions.IgnoreCase Or RegexOptions.Compiled).Groups(1).Value
                    elP.TabIndx = -1
                    lst.Add(elP)
                    'suponemos que el padre SÍ ya tiene TabIndex (por ejemplo el XtraTabcontrol)
                End If
            Next

            'tenemos rollo Nombre=Label1, TabIndx=5, NombrePadre=Panel1

            'ordenamos ya por TabIndx, muchos repetidos seguramente. 1111, 2222, 3333
            lst = lst.OrderBy(Function(x) x.TabIndx).ToList
            Dim frm As New frmTabulacions
            NameScope.SetNameScope(frm, New NameScope)
            Dim itemME = New Windows.Controls.TreeViewItem()
            itemME.Header = "Formulario"
            itemME.Name = "Me"
            frm.ArbolTabs.Items.Add(itemME)


            'añadimos sólo los padres al itemMe
            For Each sPadre In (From el As ElementoTab In lst Select el.NombrePadre Distinct).ToArray
                If sPadre = "Me" Then Continue For
                Dim elm As ElementoTab = lst.FirstOrDefault(Function(x) x.Nombre = sPadre)
                'por cada padre...
                Dim itemPadre As Windows.Controls.TreeViewItem = itemME.FindName(sPadre.Replace(".", "_"))
                If itemPadre Is Nothing Then
                    itemPadre = New Windows.Controls.TreeViewItem
                    itemPadre.Name = sPadre.Replace(".", "_")
                    itemPadre.Header = sPadre
                    itemPadre.Tag = elm
                    itemME.Items.Add(itemPadre)
                    itemME.RegisterName(itemPadre.Name, itemPadre)
                End If
                'le colgamos sus hijos
                'If Regex.Match(sPadre, ".Panel\d$").Success Then
                '    'si es un Panel de un splitter, debemos agrega un item del padre y colgarlo dentro
                '    Dim spanel As String = sPadre.Substring(0, sPadre.IndexOf(".Panel"))
                '    Dim itemsplitter As Windows.Controls.TreeViewItem = itemME.FindName(spanel.Replace(".", "_"))
                '    If itemsplitter Is Nothing Then
                '        itemsplitter = New Windows.Controls.TreeViewItem()
                '        itemsplitter.Name = spanel.Replace(".", "_")
                '        itemsplitter.Header = spanel
                '        itemsplitter.Tag = ""
                '        itemME.RegisterName(itemsplitter.Name, itemsplitter)
                '        itemME.Items.Add(itemsplitter)
                '    End If
                '    itemsplitter.Items.Add(itemPadre)
                'End If
                For Each hijo As ElementoTab In lst.Where(Function(x) x.NombrePadre = sPadre)
                    Dim itemH As Windows.Controls.TreeViewItem = itemME.FindName(hijo.Nombre.Replace(".", "_"))
                    If itemH Is Nothing Then
                        itemH = New Windows.Controls.TreeViewItem()
                        itemH.Name = hijo.Nombre.Replace(".", "_")
                        itemH.Header = hijo.Nombre
                        itemH.Tag = hijo
                        itemME.RegisterName(itemH.Name, itemH)
                    Else
                        itemME.Items.Remove(itemH)
                    End If
                    itemPadre.Items.Add(itemH)
                Next
            Next

            itemME.IsExpanded = True
            AddHandler frm.Closed, Sub(send As Object, ee As EventArgs)
                                       Dim frm1 As frmTabulacions = send
                                       If frm1.bolCerradoOK Then
                                           sbReordenaTabsRecursivo(itemME.Items, scod)
                                           IO.File.WriteAllText(sarch, scod, System.Text.Encoding.UTF8)
                                       End If
                                   End Sub
            frm.Show()
            
        End If
    End Sub

    Private Sub sbReordenaTabsRecursivo(items As Controls.ItemCollection, ByRef sCodigo As String)
        For Each item As Controls.TreeViewItem In items
            Dim elem As ElementoTab = item.Tag
            If item.Items.Count > 0 Then
                sbReordenaTabsRecursivo(item.Items, sCodigo)
            End If
            'porsiaca no tocamos lo que no hayan editado (TabModificado)
            If elem Is Nothing OrElse elem.TabIndx = -1 OrElse Not elem.TabModificado Then Continue For
            Dim nombre As String = elem.Nombre
            Dim itemP As Controls.TreeViewItem = item.Parent
            sCodigo = Regex.Replace(sCodigo, "(?<=" & nombre & "\.TabIndex = )(\d+)", itemP.Items.IndexOf(item).ToString)
        Next
    End Sub

    Public Class ElementoTab
        Public Nombre As String
        Public NombrePadre As String
        Public TabIndx As Integer
        Public TabModificado As Boolean = False
    End Class
End Class

