# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program. If not, see <http://www.gnu.org/licenses/>.




Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data

# Variables globales
$script:table = @()
$script:currentDirectory = ""


# Fonction pour récupérer la structure des dossiers
function GetFolderPathTable($path) {
    $script:table = @()
    $firstRow = @("Dossier")
    $script:table += ,$firstRow

    # Récupérer tous les dossiers de manière récursive
    $folders = Get-ChildItem -Path $path -Directory -Recurse | Sort-Object FullName

    foreach ($folder in $folders) {
        $folderPath = $folder.FullName.Substring($path.Length + 1)
        $script:table += ,@($folderPath)
    }
    return $script:table
}


#Fonction pour sauvegarder
function SaveToFile($filePath) {
    $exportTable = @()
    $exportTable += "Chemin du dossier ouvert: $script:currentDirectory"
    $headers = $script:table[0]
    $exportTable += $headers -join ','
    
    for ($i = 1; $i -lt $script:table.Count; $i++) {
        $exportTable += ($script:table[$i] -join ',')
    }

    $exportTable -join "`r`n" | Set-Content -Path $filePath
}



#Fonction de pris en compte des modifications du tableau
function UpdateTableFromDataGridView($dataGridView) {
    for ($rowIndex = 0; $rowIndex -lt $dataGridView.Rows.Count; $rowIndex++) {
        $row = $dataGridView.Rows[$rowIndex]
        for ($colIndex = 0; $colIndex -lt $row.Cells.Count; $colIndex++) {
            $script:table[$rowIndex + 1][$colIndex] = $row.Cells[$colIndex].Value
        }
    }
}



# Fonction pour charger les données depuis un fichier CSV
function LoadFromFile($filePath) {
    # Lire le contenu du fichier CSV
    $content = Get-Content -Path $filePath

    # Extraire le chemin du dossier du fichier CSV
    $script:currentDirectory = $content[0].Replace("Chemin du dossier ouvert: ", "").Trim()

    # Récupérer la liste actuelle des dossiers et leurs ACLs
    $script:table = GetFolderPathTable($script:currentDirectory)
    $script:table = ProcessAcl($script:currentDirectory)

    # Convertir le contenu du fichier CSV en une table
    $loadedTable = @()
    $rows = $content | Select-Object -Skip 1
    foreach ($row in $rows) {
        $loadedTable += ,($row -split ',')
    }

    # Comparer les données du fichier CSV avec la liste actuelle des dossiers et mettre à jour les ACLs
    for ($i = 1; $i -lt $script:table.Count; $i++) {
        $folderPath = $script:table[$i][0]
        $csvRow = $loadedTable | Where-Object { $_[0] -eq $folderPath }
        if ($csvRow) {
            for ($j = 1; $j -lt $csvRow.Count; $j++) {
                $script:table[$i][$j] = $csvRow[$j]
            }
        }
    }
}





# Fonction pour traiter les ACLs
function ProcessAcl($path) {
    $folders = Get-ChildItem -Path $path -Directory -Recurse

    foreach ($folder in $folders) {
        $folderPath = $folder.FullName.Substring($path.Length + 1)
        $acl = Get-Acl -Path (Join-Path $path $folderPath)

        # Trouver l'index de la ligne
        $rowIndex = $script:table.IndexOf(($script:table | Where-Object { $_[0] -eq $folderPath }))

        if ($rowIndex -eq -1) { continue }  # Si l'index n'est pas trouvé, passez à la prochaine itération.

        foreach ($access in $acl.Access) {
            $identityReference = $access.IdentityReference.Value
            if ($identityReference -notin $script:table[0]) {
                $script:table[0] += $identityReference
                for ($i = 1; $i -lt $script:table.Count; $i++) {
                    $script:table[$i] += "No Access"
                }
            }

            $columnIndex = $script:table[0].IndexOf($identityReference)
            
            $aclRights = $access.FileSystemRights
            if ($aclRights.HasFlag([System.Security.AccessControl.FileSystemRights]::FullControl)) {
                $script:table[$rowIndex][$columnIndex] = "FullControl"
            }
            elseif ($aclRights.HasFlag([System.Security.AccessControl.FileSystemRights]::Modify)) {
                $script:table[$rowIndex][$columnIndex] = "Modify"
            }
            elseif ($aclRights.HasFlag([System.Security.AccessControl.FileSystemRights]::ReadAndExecute)) {
                $script:table[$rowIndex][$columnIndex] = "ReadAndExecute"
            }
            else {
                $script:table[$rowIndex][$columnIndex] = "Aucun"
            }
        }
    }

    return $script:table
}

# Fonction pour mettre à jour la vue DataGridView
function UpdateDataGridView($dataGridView) {
    $dataTable = New-Object System.Data.DataTable
    $headers = $script:table[0]
    foreach ($header in $headers) {
        $dataTable.Columns.Add($header)
    }

    for ($i = 1; $i -lt $script:table.Count; $i++) {
        $dataTable.Rows.Add($script:table[$i])
    }

    $dataGridView.DataSource = $dataTable

    # Ajouter les ComboBox aux colonnes (sauf la première colonne "Dossier")
    for ($colIndex = 1; $colIndex -lt $dataGridView.ColumnCount; $colIndex++) {
        $combo = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
        $combo.Items.AddRange("FullControl", "ReadAndExecute", "Modify", "No Access")
        $combo.Name = $dataGridView.Columns[$colIndex].Name
        $combo.HeaderText = $dataGridView.Columns[$colIndex].HeaderText
        $combo.DataPropertyName = $dataGridView.Columns[$colIndex].Name  # Important pour lier les données
        $dataGridView.Columns.RemoveAt($colIndex)
        $dataGridView.Columns.Insert($colIndex, $combo)
    }
}






# GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Permissions de dossier'
$form.Size = New-Object System.Drawing.Size(1000, 600)

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Dock = [System.Windows.Forms.DockStyle]::Fill
$dataGridView.AllowUserToAddRows = $false

$dataGridView.Add_DataError({
    $args = $_
    $args.ThrowException = $false  
    $args.Cancel = $true  
})

$form.Controls.Add($dataGridView)

$openFileDialog = New-Object System.Windows.Forms.FolderBrowserDialog

# Menu fichier
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$fileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenuItem.Text = "Fichier"

# Bouton Enregistrer sous et code associé
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
# (rest of the saveFileDialog setup)
$saveAsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
# (rest of the saveAsMenuItem setup)

# Bouton ouvrir et code associé
$openMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$openMenuItem.Text = "Ouvrir"
$openMenuItem.Add_Click({
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $script:currentDirectory = $openFileDialog.SelectedPath
        $script:table = GetFolderPathTable($script:currentDirectory)
        $script:table = ProcessAcl($script:currentDirectory)
        UpdateDataGridView($dataGridView)
    }
})

# Bouton Enregistrer sous et code associé
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All files (*.*)|*.*"
$saveAsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$saveAsMenuItem.Text = "Enregistrer sous"
$saveAsMenuItem.Add_Click({
    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        UpdateTableFromDataGridView($dataGridView)
        SaveToFile($saveFileDialog.FileName)
    }
})

# Code pour le menu "Charger Sauvegarde"
$loadFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$loadFileDialog.Filter = "CSV Files (*.csv)|*.csv|All files (*.*)|*.*"
$loadMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$loadMenuItem.Text = "Charger Sauvegarde"
$loadMenuItem.Add_Click({
    if ($loadFileDialog.ShowDialog() -eq 'OK') {
        LoadFromFile($loadFileDialog.FileName)
        UpdateDataGridView($dataGridView)
    }
})

# Créez un élément de menu "Quitter" sous le menu "Fichier"
$exitMenuItem = New-Object Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Quitter"
$exitMenuItem.Add_Click({
    # Ajoutez ici le code pour quitter l'application
    $form.Close()
})

# Ajouter tous les sous-menus à fileMenuItem
$fileMenuItem.DropDownItems.Add($openMenuItem)
$fileMenuItem.DropDownItems.Add($saveAsMenuItem)
$fileMenuItem.DropDownItems.Add($loadMenuItem)  # On ajoute le menu "Charger Sauvegarde" ici
$fileMenuItem.DropDownItems.Add($exitMenuItem)  # Ajoutez l'option "Quitter" au menu "Fichier"

# Ajouter fileMenuItem au menuStrip
$menuStrip.Items.Add($fileMenuItem)
$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

$form.ShowDialog()




# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le
# modifier selon les termes de la GNU Affero General Public License telle que
# publiée par la Free Software Foundation, soit la version 3 de la Licence,
# soit (à votre choix) toute version ultérieure.
#
# Ce programme est distribué dans l'espoir qu'il sera utile, mais SANS AUCUNE
# GARANTIE ; sans même la garantie implicite de QUALITÉ MARCHANDE ou
# D'ADÉQUATION À UN USAGE PARTICULIER. Voir la GNU Affero General Public License
# pour plus de détails.
#
# Vous devriez avoir reçu une copie de la GNU Affero General Public License
# avec ce programme. Si ce n'est pas le cas, voir <http://www.gnu.org/licenses/>.
