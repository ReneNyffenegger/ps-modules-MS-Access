set-strictMode -version latest

function init {

   $src = get-content -raw "$psScriptRoot/SQL.cs"
   add-type -typeDef $src
}

function get-msAccess {

   try {
   #
   # See module COM ( https://renenyffenegger.ch/notes/Windows/PowerShell/modules/personal/COM )
   #
     $acc = get-activeObject access.application
   }
   catch {
     write-host $_
   }

   if ($acc -eq $null) {
      write-host '$acc is null!'
      return
   }

   return $acc
}

function show-msAccess {
   appActivate msaccess
}

function invoke-msAccessQuery {

   param (
     [parameter(mandatory=$true )][string   ] $sqlStmt_or_fileName,
     [parameter(mandatory=$false)][hashtable] $columnWidths
   )


   if (test-path -pathType leaf $sqlStmt_or_fileName) {
      $sqlStatement = get-content $sqlStmt_or_fileName -raw
   }
   else {
      $sqlStatement = $sqlStmt_or_fileName
   }

   $acc = get-msAccess
   if ($acc -eq $null) {
      write-host "acc is null"
      return
   }

   $sqlStatement = [tq84.SQL]::removeComments($sqlStatement)

   $queryName = 'tq84Query'

   $qry = $null
   try {
      $qry = get-COMPropertyValue $acc.currentDB().queryDefs  $queryName
   }
   catch [System.Runtime.InteropServices.COMException] {

     'invoke-msAccessQuery, exception getting query def'
      $_.Exception.GetType().FullName

         $_           | select *
         $_.exception | select *
         throw $_
   }
   if ($qry -ne $null) {

      $acQuery  = 1
      $acSaveNo = 2
    #
    # A queryDef can only be deleted if it is closed.
    #
      $acc.doCmd.close($acQuery, $queryName, $acSaveNo)
      $acc.currentDb().queryDefs.delete($queryName)
   }

   $qry = $acc.currentDB().createQueryDef($queryName, $sqlStatement)

   if ($columnWidths) {
      foreach ($fld in $qry.fields) {

         if ( ($colWidth = $columnWidths[$fld.name]) -ne $null) {
            write-host "  colWidth for $($fld.name) = $colWidth"

            try {
              $dao_dataTypeEnum_dbInteger = 3 # DAO.DataTypeEnum.dbInteger
              $prop = $fld.createProperty('ColumnWidth', $dao_dataTypeEnum_dbInteger, $colWidth)
              $fld.properties.append($prop)
            }
            catch [System.Runtime.InteropServices.COMException] {
              'Exception setting column width'
               $_.exception | select-object *
            }
            catch {
              'Exception setting column width'
               $_.exception.getType().FullName
            }
         }
      }
   }


   $acc.doCmd.openQuery($queryName)

   show-msAccess
}

init

new-alias acc-query         invoke-msAccessQuery
