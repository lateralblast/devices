param (
  [switch] $help,
  [switch] $version,
  [switch] $verbose,
  [switch] $longrackname,
  [string] $inputfile, 
  [string] $outputfile
)

# Name:         Devices
# Version:      0.2.4
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

# Import module

Import-Module VisioBot3000 -Force

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
  if (!($output_file -match ":")) {
    $output_file = "$script_dir\$output_file"
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
$dell_rack_stencils_file           = "$stencil_dir\dell\Dell-Racks.vss"
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

# Default Rack

$default_rack = "4220 Rack Frame"

# Text defaults

$default_text_size   = "12pt"
$default_text_colour = "RGB(255, 165, 0)"

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

$visio = New-VisioApplication

if ($input_file -match "csv$") {
  # Open Visio Document
  $new_doc = New-VisioDocument $output_file
  foreach ($rack in $csv_racks) {
    $cur_rows = $csv_rows | Where {$_.Rack -eq "$rack"}
    if ($longrackname) {
      $list  = @()
      $items = $cur_rows | Where {$_.Component -match "CH|Chassis"}
      $max_items = $items.count
      if (!($max_items -match "[0-9]")) {
         $max_items = 1
      }
      for ($num = 0; $num -lt $max_items; $num++) {
        $item = $cur_rows[$num].hostname
        $list += $item
      }
      $hosts     = $list -join ","
      $rack_name = "$rack ($hosts)"
    }
    else {
      $rack_name = "$rack"
    }
    $new_page = New-VisioPage -Name $rack_name
  }
  $remove_page = Remove-VisioPage -Name "Page-1"
  # Setup Rack Stencils
  # Dell has a good default rack stencil
  $dell_rack_stencils = Register-VisioStencil -Name dell_rack_stencils -Path $dell_rack_stencils_file 
  $rack_stencil       = Register-VisioShape -Name rack_stencil -From dell_rack_stencils -MasterName "$default_rack"
  # Check for vendors
  $pure_rows = $csv_rows | Where {$_.Vendor -match "Pure"}
  if ($pure_rows) {
    $pure_storage_array_stencils = Register-VisioStencil -Name pure_storage_array_stencils -Path $pure_storage_array_stencils_file 
  }
  $dell_rows = $csv_rows | Where {$_.Vendor -match "Dell"}
  if ($dell_rows) {
    $model_test = $dell_rows | Where {$_.Model -match "^CX4|^NX4|^ES|^DD"}
    if ($model_test -match "[A-Z]") {
      $dell_emc_storage_stencils = Register-VisioStencil -Name dell_emc_storage_stencils  $dell_emc_storage_stencils_file
    }
    $model_test = $dell_rows | Where {$_.Model -match "^R|^C"}
    if ($model_test -match "[A-Z]") {
      $dell_rack_server_stencils = Register-VisioStencil -Name dell_rack_server_stencils $dell_rack_server_stencils_file
    }
    $model_test = $dell_rows | Where {$_.Model -match "^M[0-9]"}
    if ($model_test -match "[A-Z]") {
      $dell_blade_server_stencils = Register-VisioStencil -Name dell_blade_server_stencils $dell_blade_server_stencils_file
    }
    $model_test = $dell_rows | Where {$_.Model -match "^FS8|^SC"}
    if ($model_test -match "[A-Z]") {
      $dell_sc_storage_stencils = Register-VisioStencil -Name dell_sc_storage_stencils $dell_sc_storage_stencils_file
    }
    $model_test = $dell_rows | Where {$_.Model -match "^FS7|^PS"}
    if ($model_test -match "[A-Z]") {
      $dell_ps_storage_stencils = Register-VisioStencil -Name dell_ps_storage_stencils $dell_ps_storage_stencils_file
    }
    $model_test = $dell_rows | Where {$_.Model -match "^D|^MD|^NX"}
    if ($model_test -match "[A-Z]") {
      $dell_md_storage_stencils = Register-VisioStencil -Name dell_md_storage_stencils $dell_md_storage_stencils_file
    }
  }
  $sun_rows = $csv_rows | Where {$_.Vendor -match "Oracle|Sun"}
  if ($sun_rows) {
    $model_test = $sun_rows | Where {$_.Model -match "Blade|^B[0-9]"}
    if ($model_test -match "[A-Z]") {
      $oracle_blade_server_stencils = Register-VisioStencil -Name oracle_blade_server_stencils $oracle_blade_server_stencils_file 
    }
    $model_test = $sun_rows | Where {$_.Model -match "SPARC|sparc|^T[0-9]|^M[0-9]|^E[0-9]"}
    if ($model_test -match "[A-Z]") {
      $oracle_sparc_server_stencils = Register-VisioStencil -Name oracle_sparc_server_stencils $oracle_sparc_server_stencils_file 
    }
    $model_test = $sun_rows | Where {$_.Model -match "X64|X86|x64|x86|i386|^X[0-9]"}
    if ($model_test -match "[A-Z]") {
      $oracle_intel_server_stencils = Register-VisioStencil -Name oracle_intel_server_stencils $oracle_intel_server_stencils_file 
    }
  }
  foreach ($rack in $csv_racks) {
    # Reset grid references
    $cur_front_ru_x = $front_ru_x 
    $cur_back_ru_x  = $back_ru_x 
    $cur_front_ru_y = $front_ru_y 
    $cur_back_ru_y  = $back_ru_y 
    # Process rows in CSV for current rack
    $cur_rows = $csv_rows | Where {$_.Rack -eq "$rack"}
    $max_rows = $cur_rows.count
    if (!($max_rows -match "[0-9]")) {
       $max_rows = 1
    }
    if ($longrackname) {
      $list  = @()
      $items = $cur_rows | Where {$_.Component -match "CH|Chassis"}
      $max_items = $items.count
      if (!($max_items -match "[0-9]")) {
         $max_items = 1
      }
      for ($num = 0; $num -lt $max_items; $num++) {
        $item = $cur_rows[$num].hostname
        $list += $item
      }
      $hosts     = $list -join ","
      $rack_name = "$rack ($hosts)"
    }
    else {
      $rack_name = "$rack"
    }
    # Select Rack Page
    $page = Set-VisioPage $rack_name
    # Place rack front stencil
    $location   = Set-NextShapePosition -x $front_rack_x -y $front_rack_y
    $shape      = rack_stencil rack_front
    $label      = $shape.Characters.Text=$rack_name
    $colour     = $shape.Cells("Char.Color").FormulaU = $default_text_colour
    $colour     = $shape.Cells("Char.Size").FormulaU  = $default_text_size
    $shape_data = Set-VisioShapeData -Shape $rack_front -Name ProductNumber ""
    $shape_data = Set-VisioShapeData -Shape $rack_front -Name Manufacturer ""
    $shape_data = Set-VisioShapeData -Shape $rack_front -Name ProductNumber ""
    $shape_data = Set-VisioShapeData -Shape $rack_front -Name PartNumber ""
    $shape_data = Set-VisioShapeData -Shape $rack_front -Name ProductDescription ""
    # Place rack back stencil
    $location   = Set-NextShapePosition -x $back_rack_x -y $back_rack_y
    $shape      = rack_stencil rack_back
    $label      = $shape.Characters.Text=$rack_name
    $colour     = $shape.Cells("Char.Color").FormulaU = $default_text_colour
    $colour     = $shape.Cells("Char.Size").FormulaU  = $default_text_size
    $shape_data = Set-VisioShapeData -Shape $rack_back -Name ProductNumber ""
    $shape_data = Set-VisioShapeData -Shape $rack_back -Name Manufacturer ""
    $shape_data = Set-VisioShapeData -Shape $rack_back -Name ProductNumber ""
    $shape_data = Set-VisioShapeData -Shape $rack_back -Name PartNumber ""
    $shape_data = Set-VisioShapeData -Shape $rack_back -Name ProductDescription ""
    # Set through rows on current rack
    for ($row = 0; $row -lt $max_rows; $row++) {
      # Get Values from columns
      $hostname  = $cur_rows[$row].Hostname
      $component = $cur_rows[$row].Component
      $vendor    = $cur_rows[$row].Vendor
      $arch      = $cur_rows[$row].Architecture
      $model     = $cur_rows[$row].Model
      $os        = $cur_rows[$row]."Operating System"
      $rack      = $cur_rows[$row].Rack
      $rus       = $cur_rows[$row]."Rack Units"
      $top_ru    = $cur_rows[$row]."Top Rack Unit"
      $serial    = $cur_rows[$row]."Serial Number"
      $asset     = $cur_rows[$row]."Asset Number"
      $installed = $cur_rows[$row]."Installed Date"
      $warranty  = $cur_rows[$row]."Warranty Exp"
      $location  = $cur_rows[$row].Location
      $country   = $cur_rows[$row].Country
      $info      = "$hostname"+": $component"
      # Handle Vendor to help chose shapes
      switch -regex ($vendor) {
        # Handle Dell Servers, Blades and Storage
        "Dell" {
          $front_name = "$model Front"
          $back_name  = "$model Rear"
          switch -regex ($model) {
            "^CX4|^NX4|^ES|^DD" {
              $dell_emc_storage_stencil_front = Register-VisioShape -Name stencil_front -From dell_emc_storage_stencils -MasterName "$front_name"
              $dell_emc_storage_stencil_back  = Register-VisioShape -Name stencil_back  -From dell_emc_storage_stencils -MasterName "$back_name"
            }
            "^R|^C" {
              $dell_rack_server_stencil_front = Register-VisioShape -Name stencil_front -From dell_rack_server_stencils -MasterName "$front_name"
              $dell_rack_server_stencil_back  = Register-VisioShape -Name stencil_back  -From dell_rack_server_stencils -MasterName "$back_name"
            }
            "^M[0-9]" {
              $dell_blade_server_stencil_front = Register-VisioShape -Name stencil_front -From dell_blade_server_stencils -MasterName "$front_name"
              $dell_blade_server_stencil_back  = Register-VisioShape -Name stencil_back  -From dell_blade_server_stencils -MasterName "$back_name"
            }
            "^FS8|^SC" {
              $dell_sc_storage_stencil_front = Register-VisioShape -Name stencil_front -From dell_sc_storage_stencils -MasterName "$front_name"
              $dell_sc_storage_stencil_back  = Register-VisioShape -Name stencil_back  -From dell_sc_storage_stencils -MasterName "$back_name"
            }
            "^FS7|^PS" {
              $dell_ps_storage_stencil_front = Register-VisioShape -Name stencil_front -From dell_ps_storage_stencils -MasterName "$front_name"
              $dell_ps_storage_stencil_back  = Register-VisioShape -Name stencil_back  -From dell_ps_storage_stencils -MasterName "$back_name"
            }
            "^D|^MD|^NX" {
              $dell_md_storage_stencil_front = Register-VisioShape -Name stencil_front -From dell_md_storage_stencils -MasterName "$front_name"
              $dell_md_storage_stencil_back  = Register-VisioShape -Name stencil_back  -From dell_md_storage_stencils -MasterName "$back_name"
            }
          }
        }
        # Handle Pure arrays
        "Pure" {
          switch -regex ($model) {
            "FB|FlashBlade" {
              $front_name = "FlashBlade Front Full"
              $back_name  = "FlashBlade back"
            }
            "FA-|M" {
              $front_name = "FA M70 front"
              $back_name  = "FA M70 back"
            }
            default {
              $front_name = "$model front"
              $back_name  = "$model back"
            }
          }
          $pure_storage_array_stencil_front = Register-VisioShape -Name stencil_front -From pure_storage_array_stencils -MasterName "$front_name"
          $pure_storage_array_stencil_back  = Register-VisioShape -Name stencil_back  -From pure_storage_array_stencils -MasterName "$back_name"
        }
        # Handle Oracle servers
        "Oracle|Sun" {
          $front_name = "$model Front" 
          $back_name  = "$model Rear"
          if ($model -match "Blade") {
            $oracle_blade_server_stencil_front = Register-VisioShape -Name stencil_front -From oracle_blade_server_stencils -MasterName "$front_name"
            $oracle_blade_server_stencil_back  = Register-VisioShape -Name stencil_back  -From oracle_blade_server_stencils -MasterName "$back_name"
          }
          else {
            if ($arch -match "SPARC|sparc") {
              $oracle_sparc_server_stencil_front = Register-VisioShape -Name stencil_front -From oracle_sparc_server_stencils -MasterName "$front_name"
              $oracle_sparc_server_stencil_back  = Register-VisioShape -Name stencil_back  -From oracle_sparc_server_stencils -MasterName "$back_name"
            }
            else {
              $oracle_intel_server_stencil_front = Register-VisioShape -Name stencil_front -From oracle_intel_server_stencils -MasterName "$front_name"
              $oracle_intel_server_stencil_back  = Register-VisioShape -Name stencil_back  -From oracle_intel_server_stencils -MasterName "$back_name"
            }
          }
        }
        default {
          $front_name           = "1U Metal Close Out"
          $back_name            = "1U Metal Close Out"
          $blank_stencil_front  = Register-VisioShape -Name stencil_front -From dell_rack_stencils -MasterName "$front_name"
          $blank_stencil_back   = Register-VisioShape -Name stencil_back  -From dell_rack_stencils -MasterName "$back_name"
        }
      }
      # Place stencils
      $rack_space     = [float]$rus * [float]$ru_space
      $cur_front_ru_y = [float]$front_ru_y + ([float]$top_ru * [float]$ru_space) - $rack_space
      $cur_back_ru_y  = [float]$back_ru_y + ([float]$top_ru * [float]$ru_space) - $rack_space
      $location_x     = Set-NextShapePosition -x $cur_front_ru_x -y $cur_front_ru_y
      $shape          = stencil_front stencil
      $shape_label    = $shape.Characters.Text=$info
      $text_colour    = $shape.Cells("Char.Color").FormulaU = $default_text_colour
      $text_size      = $shape.Cells("Char.Size").FormulaU  = $default_text_size
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name SerialNumber $serial
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name AssetNumber $asset
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name Location $location
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name InstalledDate $installed
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name WarrantyExp $warranty
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name Room $rack
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name RackUnits $rus
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name OperatingSystem $os
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name SystemName $hostname
      $location_x     = Set-NextShapePosition -x $cur_back_ru_x -y $cur_back_ru_y
      $shape          = stencil_back stencil
      $shape_label    = $shape.Characters.Text=$info
      $text_colour    = $shape.Cells("Char.Color").FormulaU = $default_text_colour
      $text_size      = $shape.Cells("Char.Size").FormulaU  = $default_text_size
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name SerialNumber $serial
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name AssetNumber $asset
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name Location $location
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name InstalledDate $installed
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name WarrantyExp $warranty
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name Room $rack
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name RackUnits $rus
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name OperatingSystem $os
      $shape_data     = Set-VisioShapeData -Shape $stencil -Name SystemName $hostname
    }
  }
  $doc = Complete-VisioDocument -Close
}
