param (
  [switch] $help,
  [switch] $version,
  [switch] $verbose,
  [string] $inputfile  = "servers.xlsx", 
  [string] $outputfile = "servers.vsd"
)

# Name:         vole (Visio diagram Output using Locations from Excel) 
# Version:      0.1.7
# Release:      1
# License:      CC-BA (Creative Commons By Attribution)
#               http://creativecommons.org/licenses/by/4.0/legalcode
# Group:        System
# Source:       N/A
# URL:          http://lateralblast.com.au/
# Distribution: UNIX
# Vendor:       Lateral Blast
# Packager:     Richard Spindler <richard@lateralblast.com.au>
# Description:  Powershell script to output a Visio diagram from an Excel file

# Script glabal vars

$script_name = $MyInvocation.MyCommand.Name
$script_path = $MyInvocation.MyCommand.Path
$script_file = $MyInvocation.MyCommand
$script_dir  = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$script_vers = ""
$data_dir    = "$script_dir\data"
$stencil_dir = "$script_dir\stencils"
$output_dir  = "$script_dir\output"

$script_text = Get-Content $script_file

function unzip_stencil($stencil_file) {
  $zip_file    = "$stencil_file.zip"
  $shell_obj   = new-object -com shell.application
  $zip_obj     = $shell_obj.NameSpace($zip_file)
  $destination = Split-Path $stencil_file
  if (!(Test-Path $stencil_file)) {
    foreach($item in $zip_obj.items()) {
      Write-Host "Extracting '$item' from '$zip_file' to '$destination'"
      $shell_obj.Namespace($destination).copyhere($item)
    }
  }
  return
}

function get_file_type($input_file) {
  Add-Type -AssemblyName "System.Web"
  $mime_type = [System.Web.MimeMapping]::GetMimeMapping($script_file)
  return($mime_type)
}

function print_help($script_name) {
  Write-Host "usage: $script_name"
  Write-Host "--help"
  Write-Host "--version"
  Write-Host "--inputfile  FILENAME"
  Write-Host "--outputfile FILENAME"
  return
}

function print_version($script_vers) {
  Write-Host "$script_vers"
}

function get_script_vers($script_text) {
  foreach ($script_line in $script_text) {
    if ($script_line -match "# Version") {
      $line_items  = $script_line -split "\s+"
      $script_vers = $line_items[2]
      return $script_vers
    }
  }
}

$script_vers = get_script_vers($script_text)

# Print help

if ($help) {
  print_help($script_name)
  exit
}

# Print version

if ($version) {
  print_version($script_vers)
  exit
}

# Handle input switch

if ($inputfile) {
  $input_file = $inputfile
  if ($input_file -match "^servers.xlsx$") {
    $input_file =  "$data_dir\servers.xlsx"
  }
  if (!(Test-Path $input_file)) {
    Write-Host "File: '$input_file' does not exist"
    exit
  }
}
else {
  Write-Host "Input file not specified"
  print_help($script_name)
  exit
}

if ($verbose) {
  Write-Host "Input:  $input_file"
}

# Handle output switch

if ($outputfile) {
  $output_file = $outputfile
  if ($output_file -match "^servers.vsd$") {
    $output_file =  "$output_dir\servers.vsd"
  }
}
else {
  Write-Host "Output file not specified"
  print_help($script_name)
  exit
}

# Get file type

$file_type = get_file_type($input_file)

# Handle opening file

if ($input_file -match "xls$|xlsx$" -And $file_type -match "octet") {
  $excel   = New-Object -ComObject Excel.Application
  $book    = $excel.Workbooks.Open($input_file)
  $sheet   = $book.Worksheets.Item(1)
  $max_row = ($sheet.UsedRange.Rows).count 
  $max_col = ($sheet.UsedRange.Columns).count
}
else {
  if ($input_file -match "csv$") {
    $csv_rows = Import-Csv $input_file
  }
}

$csv_racks = $csv_rows.Rack | Select-Object -Unique

# Set up some global stencil files

$oracle_sparc_server_stencils_file = "$stencil_dir\oracle\Oracle-Server-SPARC.vss"
$oracle_intel_server_stencils_file = "$stencil_dir\oracle\Oracle-Server-x86.vss"
$oracle_blade_server_stencils_file = "$stencil_dir\oracle\Oracle-Server-Blade.vss"
$dell_rack_stencils_file 	         = "$stencil_dir\dell\Dell-Racks.vss"
$dell_blade_server_stencils_file   = "$stencil_dir\dell\Dell-PowerEdge-BladeServers.vss"
$dell_rack_server_stencils_file    = "$stencil_dir\dell\Dell-PowerEdge-RackServers.vss"
$dell_sc_storage_stencils_file     = "$stencil_dir\dell\Dell-Storage-Compellent-SC.vss"
$dell_ps_storage_stencils_file     = "$stencil_dir\dell\Dell-Storage-EqualLogic-PS.vss"
$dell_md_storage_stencils_file     = "$stencil_dir\dell\Dell-Storage-PowerVault-Dx-MD-NX.vss"
$dell_emc_storage_stencils_file    = "$stencil_dir\dell\Dell-EMC.vss"
$ibm_power_stencils_file           = "$stencil_dir\ibm\IBM-Server-Power.vss"
$ibm_systemi_stencils_file         = "$stencil_dir\ibm\IBM-Server-Systemi.vss"
$ibm_systemp_stencils_file         = "$stencil_dir\ibm\IBM-Server-Systemp.vss"
$ibm_systemx_stencils_file         = "$stencil_dir\ibm\IBM-Server-Systemx.vss"
$ibm_systemz_stencils_file         = "$stencil_dir\ibm\IBM-Server-Systemz.vss"
$pure_storage_array_stencils_file  = "$stencil_dir\pure\Purestorage.vss"
$netapp_nearstore_stencils_file    = "$stencil_dir\netapp\NetApp-NearStore-classic.vss"
$netapp_fas_stencils_file          = "$stencil_dir\netapp\NetApp-FAS-Series.vss"
$netapp_old_fas_stencils_file      = "$stencil_dir\netapp\NetApp-FAS-Series-classic.vss"
$netapp_e_series_stencils_file     = "$stencil_dir\netapp\NetApp-E-Series.vss"
$netapp_s_series_stencils_file     = "$stencil_dir\netapp\NetApp-S-Family-classic.vss"
$netapp_v_series_stencils_file     = "$stencil_dir\netapp\NetApp-V-Series.vss"
$netapp_old_v_series_stencils_file = "$stencil_dir\netapp\NetApp-V-Series-classic.vss"
$netapp_vtl_stencils_file          = "$stencil_dir\netapp\NetApp-VTL-Series-classic.vss"

# Rack x,y constants

$front_rack_x   = 1.0
$back_rack_x    = 5.0
$front_rack_y   = 4.0
$back_rack_y    = 4.0
$front_ru_x     = $front_rack_x + 1.19
$back_ru_x      = $back_rack_x + 1.19
$front_ru_y     = 2.13
$back_ru_y      = 2.13
$ru_space       = 0.175
$cur_rack       = "None"

$visio = New-Object -ComObject Visio.Application

if ($input_file -match "csv$") {
  foreach ($rack in $csv_racks) {
    $output_file =  "$output_dir\rack-$rack.vsd"
    # Open Visio Document
    $doc   = $visio.Documents.Add("")
    $page  = $visio.ActiveDocument.Pages.Item(1)
    # Dell has a good default rack stencil
    unzip_stencil($dell_rack_stencils_file)
    $dell_rack_stencils = $visio.Documents.Add($dell_rack_stencils_file)
    $rack_stencil       = $dell_rack_stencils.Masters.Item("4220 Rack Frame")
    # Place rack front stencil
    $front_rack_shape = $page.Drop($rack_stencil,$front_rack_x,$front_rack_y)
    # Place rack front stencil
    $back_rack_shape = $page.Drop($rack_stencil,$back_rack_x,$back_rack_y)
    $cur_rows = $csv_rows | Where {$_.Rack -eq "$rack"}
    $max_rows = $cur_rows.count
    if (!($max_rows -match "[0-9]")) {
      $max_rows = 1
    }
    for ($row = 0; $row -lt $max_rows; $row++) {
      # Get Values from columns
      $hostname = $csv_rows[$row].Hostname
      $model    = $csv_rows[$row].Model
      $top_ru   = $csv_rows[$row]."Top Rack Unit"
      $vendor   = $csv_rows[$row].Vendor
      $rack     = $csv_rows[$row].Rack
      $rus      = $csv_rows[$row]."Rack Units"
      $arch     = $csv_rows[$row].Architecture
      # Calculate bottom RU locations
      if ($row -eq 0) {
        $cur_front_ru_x = $front_ru_x 
        $cur_back_ru_x  = $back_ru_x 
        $cur_front_ru_y = $front_ru_y 
        $cur_back_ru_y  = $back_ru_y 
      }
      else {
        $cur_front_ru_x = $front_ru_x 
        $cur_back_ru_x  = $back_ru_x 
      }
      $server_stencils = ""
      switch -regex ($vendor) {
        # Handle NetApp Storage
        "NetApp" {
          $front_name = "$model Front"
          $back_name  = "$model rear"
          switch -regex ($model) {
            "^E|^DE" {
              if (!($netapp_e_series_stencils)) {
                unzip_stencil($netapp_e_series_stencils_file)
                $netapp_e_series_stencils = $visio.Documents.Add($netapp_e_series_stencils_file)
              }
              $server_stencils = $netapp_e_series_stencils
            }
            "^R" {
              if (!($netapp_nearstore_stencils)) {
                unzip_stencil($netapp_nearstore_stencils_file)
                $netapp_nearstore_stencils = $visio.Documents.Add($netapp_nearstore_stencils_file)
              }
              $server_stencils = $netapp_nearstore_stencils
            }
            "^S" {
              if (!($netapp_s_series_stencils)) {
                unzip_stencil($netapp_s_series_stencils_file)
                $netapp_s_series_stencils = $visio.Documents.Add($netapp_s_series_stencils_file)
              }
              $server_stencils = $netapp_s_series_stencils
            }
            # FAS and V series needs tweaking for model numbers
            "^FAS[82,22,32,25,26,62,90]" {
              if (!($netapp_old_fas_stencils)) {
                unzip_stencil($netapp_old_fas_stencils_file)
                $netapp_old_fas_stencils = $visio.Documents.Add($netapp_old_fas_stencils_file)
              }
              $server_stencils = $netapp_old_fas_stencils
            }
            "^FAS[0-6][0,1]|^FAS[3,6]2" {
              if (!($netapp_fas_stencils)) {
                unzip_stencil($netapp_fas_stencils_file)
                $netapp_fas_stencils = $visio.Documents.Add($netapp_fas_stencils_file)
              }
              $server_stencils = $netapp_old_fas_stencils
            }
            "^V[82,22,32,25,26,62,90]" {
              if (!($netapp__old_v_series_stencils)) {
                unzip_stencil(netapp_old_v_series_stencils_file)
                $netapp_old_v_series_stencils = $visio.Documents.Add($netapp_old_v_series_stencils_file)
              }
              $server_stencils = $netapp_old_v_series_stencils
            }
            "^V[0-6][0,1]|^V[3,6]2" {
              if (!($netapp_v_series_stencils)) {
                unzip_stencil($netapp_v_series_stencils_file)
                $netapp_v_series_stencils = $visio.Documents.Add($netapp_v_series_stencils_file)
              }
              $server_stencils = $netapp_old_v_series_stencils
            }
            "^VTL" {
              if (!($netapp_vtl_stencils)) {
                unzip_stencil($netapp_vtl_stencils_file)
                $netapp_vtl_stencils = $visio.Documents.Add($netapp_vtl_stencils_file)
              }
              $server_stencils = $netapp_vtl_stencils
            }
          }          
        }
        # Handle IBM Servers
        "IBM" {
          $front_name = "$model Rack Front"
          $back_name  = "$model Rack Rear"
          switch -regex ($model) {
            "Power" {
              if (!($ibm_power_stencils)) {
                unzip_stencil($ibm_power_stencils_file)
                $ibm_power_stencils = $visio.Documents.Add($ibm_power_stencils_file)
              }
              $server_stencils = $ibm_power_stencils
            }
            "^i" {
              if (!($ibm_systemi_stencils)) {
                unzip_stencil($ibm_systemi_stencils_file)
                $ibm_systemi_stencils = $visio.Documents.Add($ibm_systemi_stencils_file)
              }
              $server_stencils = $ibm_systemi_stencils
            }
            "^p" {
              if (!($ibm_systemp_stencils)) {
                unzip_stencil(ibm_systemp_stencils_file)
                $ibm_systemp_stencils = $visio.Documents.Add($ibm_systemp_stencils_file)
              }
              $server_stencils = $ibm_systemp_stencils
            }
            "^x" {
              if (!($ibm_systemx_stencils)) {
                unzip_stencil($ibm_systemx_stencils_file)
                $ibm_systemx_stencils = $visio.Documents.Add($ibm_systemx_stencils_file)
              }
              $server_stencils = $ibm_systemx_stencils
            }
            "^z" {
              if (!($ibm_systemz_stencils)) {
                unzip_stencil($ibm_systemz_stencils_file)
                $ibm_systemz_stencils = $visio.Documents.Add($ibm_systemz_stencils_file)
              }
              $server_stencils = $ibm_systemz_stencils
            }
          }          
        }
        # Handle Dell Servers, Blades and Storage
        "Dell" {
          $front_name = "$model Front"
          $back_name  = "$model Rear"
          switch -regex ($model) {
            "^CX4|^NX4|^ES|^DD" {
              if (!($dell_rack_server_stencils)) {
                unzip_stencil($dell_emc_storage_stencils_file)
                $dell_emc_storage_stencils = $visio.Documents.Add($dell_emc_storage_stencils_file)
              }
              $server_stencils = $dell_emc_storage_stencils
            }
            "^R|^C" {
              if (!($dell_rack_server_stencils)) {
                unzip_stencil($dell_rack_server_stencils_file)
                $dell_rack_server_stencils = $visio.Documents.Add($dell_rack_server_stencils_file)
              }
              $server_stencils = $dell_rack_server_stencils
            }
            "^M[0-9]" {
              if (!($dell_blade_server_stencils)) {
                unzip_stencil($dell_blade_server_stencils_file)
                $dell_blade_server_stencils = $visio.Documents.Add($dell_blade_server_stencils_file)
              }
              $server_stencils = $dell_blade_server_stencils
            }
            "^FS8|^SC" {
              if (!($dell_sc_storage_stencils)) {
                unzip_stencil($dell_sc_storage_stencils_file)
                $dell_sc_storage_stencils = $visio.Documents.Add($dell_sc_storage_stencils_file)
              }
              $server_stencils = $dell_sc_storage_stencils_file
            }
            "^FS7|^PS" {
              if (!($dell_ps_storage_stencils)) {
                unzip_stencil($dell_ps_storage_stencils_file)
                $dell_ps_storage_stencils = $visio.Documents.Add($dell_ps_storage_stencils_file)
              }
              $server_stencils = $dell_ps_storage_stencils_file
            }
            "^D|^MD|^NX" {
              if (!($dell_md_storage_stencils)) {
                unzip_stencil($dell_md_storage_stencils_file)
                $dell_md_storage_stencils = $visio.Documents.Add($dell_md_storage_stencils_file)
              }
              $server_stencils = $dell_md_storage_stencils_file
            }
          }
        }
        # Handle Pure arrays
        "Pure" {
          $front_name = "$model front"
          $back_name  = "$model back"
          if (!($pure_storage_array_stencils)) {
            unzip_stencil($pure_storage_array_stencils_file)
            $pure_storage_array_stencils = $visio.Documents.Add($pure_storage_array_stencils_file)
          }
          $server_stencils = $pure_storage_array_stencils
        }
        # Handle Oracle servers
        "Oracle|Sun" {
          $front_name = "$model Front" 
          $back_name  = "$model Rear"
          if ($model -match "Blade") {
            if (!($oracle_blade_server_stencils)) {
              unzip_stencil($oracle_blade_server_stencils_file)
              $oracle_blade_server_stencils = $visio.Documents.Add($oracle_blade_server_stencils_file)
            }
            $server_stencils = $oracle_blade_server_stencils
          }
          else {
            if ($arch -match "SPARC|sparc") {
              if (!($oracle_sparc_server_stencils)) {
                unzip_stencil($oracle_sparc_server_stencils_file)
                $oracle_sparc_server_stencils = $visio.Documents.Add($oracle_sparc_server_stencils_file)
              }
              $server_stencils = $oracle_sparc_server_stencils
            }
            else {
              if (!($oracle_intel_server_stencils)) {
                unzip_stencil($oracle_intel_server_stencils_file)
                $oracle_intel_server_stencils = $visio.Documents.Add($oracle_intel_server_stencils_file)
              }
              $server_stencils = $oracle_intel_server_stencils
            }
          }
        }
        default {
          $front_name      = "1U Metal Close Out"
          $back_name       = "1U Metal Close Out"
          $server_stencils = $dell_rack_stencils
        }
      }
      # Place stencils
      $front_stencil = $server_stencils.Masters.Item($front_name)
      $back_stencil  = $server_stencils.Masters.Item($back_name)
      $front_test    = $front_stencil.Name
      $back_test     = $back_stencil.Name
      if (!($front_test -match "$front_name")) {
        $front_name      = "1U Metal Close Out"
        $server_stencils = $dell_rack_stencils
        $front_stencil   = $server_stencils.Masters.Item($front_name)
      }
      if (!($back_test -match "$back_name")) {
        $back_name       = "1U Metal Close Out"
        $server_stencils = $dell_rack_stencils
        $back_stencil    = $server_stencils.Masters.Item($back_name)
      }
      if ($front_name -match "1U Metal Close Out") {
        for ($count = 1; $count -le [int]$rus ; $count++) {
          $front_shape    = $page.Drop($front_stencil,$cur_front_ru_x,$cur_front_ru_y)
          $rack_space     = [float]$ru_space
          $cur_front_ru_y = [float]$cur_front_ru_y + [float]$rack_space
        } 
      }
      else {
        $front_shape    = $page.Drop($front_stencil,$cur_front_ru_x,$cur_front_ru_y)
        $rack_space     = [float]$rus * [float]$ru_space
        $cur_front_ru_y = [float]$cur_front_ru_y + [float]$rack_space
      }
      if ($back_name -match "1U Metal Close Out") {
        for ($count = 1; $count -le [int]$rus ; $count++) {
          $back_shape    = $page.Drop($front_stencil,$cur_back_ru_x,$cur_back_ru_y)
          $rack_space    = [float]$ru_space
          $cur_back_ru_y = [float]$cur_back_ru_y + [float]$rack_space
        }
      }
      else {
        $back_shape    = $page.Drop($back_stencil,$cur_back_ru_x,$cur_back_ru_y)
        $rack_space    = [float]$rus * [float]$ru_space
        $cur_back_ru_y = [float]$cur_back_ru_y + [float]$rack_space
      }
    }
    # Output file  
    $doc.SaveAs($output_file)
    Write-Host "Output: $output_file"
  }
}

