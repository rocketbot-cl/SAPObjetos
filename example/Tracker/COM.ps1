
#-Begin-----------------------------------------------------------------

#-Load assembly---------------------------------------------------------
Add-Type -AssemblyName "Microsoft.VisualBasic";
Add-Type -AssemblyName "System.Windows.Forms";

#-Function Create-Object------------------------------------------------
Function Create-Object {

  Param(
    [String]$objectName
  )

  Try {
    New-Object -ComObject $objectName;
  } Catch {
    [Void][System.Windows.Forms.MessageBox]::Show(
      "Can't create object", "Important hint", 0);
  }

}

#-Function Get-Object---------------------------------------------------
Function Get-Object {

  Param(
    [String]$pathName,
    [String]$class
  )

  [Microsoft.VisualBasic.Interaction]::GetObject($pathName, $class);

}

#-Sub Free-Object-------------------------------------------------------
Function Free-Object {

  Param(
    [__ComObject]$object
  )

  [Void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($object);

}

#-Function Get-Property-------------------------------------------------
Function Get-Property {

  Param(
    [__ComObject]$object,
    [String]$propertyName, 
    $propertyParameter
  )

  $objectType = [System.Type]::GetType($object);
  $objectType.InvokeMember($propertyName,
    [System.Reflection.Bindingflags]::GetProperty,
    $null, $object, $propertyParameter);

}

#-Sub Set-Property------------------------------------------------------
Function Set-Property {

  Param(
    [__ComObject]$object,
    [String]$propertyName,
    $propertyValue
  )

  $objectType = [System.Type]::GetType($object);
  [Void] $objectType.InvokeMember($propertyName,
    [System.Reflection.Bindingflags]::SetProperty,
    $null, $object, $propertyValue);

}

#-Function Invoke-Method------------------------------------------------
Function Invoke-Method {

  Param(
    [__ComObject]$object,
    [String]$methodName,
    $methodParameters
  )

  $objectType = [System.Type]::GetType($object);
  $output = $objectType.InvokeMember($methodName,
    [System.Reflection.BindingFlags]::InvokeMethod,
    $null, $object, $methodParameters);
  if ( $output ) { $output }

}

#-End-------------------------------------------------------------------
