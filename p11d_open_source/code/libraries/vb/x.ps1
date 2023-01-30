
function full_path($path, $path_separator = "\", $front = $false)
{
  if ( $path -eq $null )
  { return $null; }

  if ( $path -is [system.array] )
  {
    $path_string = ""; 
    foreach( $path_local in $path )
    {
      if ( $path_string.Length -gt 0 )
      {
        [string] $s = (full_path $path_local $path_separator $front);

        if ( $s.Length -gt 0 -and  $s[0] -eq $path_separator )
        {
          $s = $s.Substring(1);
        }

        $path_string += $s
      } 
      else
      {
        $path_string = (full_path $path_local $path_separator $front)
      }
        
    }
    return $path_string;
  }
  

  $path = $path.Trim();
  if ( $path.Length -eq 0 )
  {
    return $path_separator;
  }
  if ( ! $front )
  {
    if ( $path[$path.Length - 1] -ne $path_separator )
    {
      $path = ($path + $path_separator);     
    }
  }
  else
  {
    if ( $path[0] -ne $path_separator )
    {
      $path = ($path_separator + $path);     
    }
  }

  return $path;
}

function file_exist($file_and_path)
{
  if ( $file_and_path -eq $null )
  { return $false; }
  
    return (Test-Path -Path $file_and_path -PathType Leaf);    
}

function directory_exist($path)
{
    if ( $path -eq $null )
    { return $false; }
    return (Test-Path -Path $path);    
}

function windows_system_path_32()
{
    $windows_system_path = [Environment]::SystemDirectory;
    $windows_path = (get-item $windows_system_path ).parent.FullName;
    
    $windows_system_path_32 = Join-Path -Path $windows_path -ChildPath "SysWow64";
    if ( Test-Path -Path $windows_system_path_32 )
    {
        $windows_system_path = $windows_system_path_32;        
    }

    return $windows_system_path
}

function working_directory()
{
    $ret = $PWD.ProviderPath -replace '\\', '/'  
    return $ret;
}


function full_file_path( $file_name )
{
    $working_dir = working_directory;
    $file_full_path_val = full_path $working_dir;
    $ret = "$($file_full_path_val)$($file_name)";
    return $ret;
}
function copy_file_to_release($file)
{
    $file_path = full_file_path $file
    if ( !(file_exist $file_path)) 
    {
      return;
    }
    
    $release_dir = (full_path $working_dir ) + "release"
    if ( ! (directory_exist $release_dir) ) 
    {
        throw "release folder $($release_dir) does not exist";
    }

    Copy-Item $file_path $release_dir -Force;
    Write-Host "Copied $($file) to $($release_dir)"
}


function copy_file_to_win_sys($file)
{
    $file_path = full_file_path $file
    if ( !(file_exist $file_path)) 
    {
      return $false;
    }
    
    $destination_path = windows_system_path_32;
    Copy-Item $file_path $destination_path -Force -errorAction stop;
    Write-Host "Copied $($file) to $($destination_path)"
    return $true
}   



function copy_file_to_win_sys_and_register($file)
{
    $destination_path = windows_system_path_32;
    $destination_path = full_path $destination_path;    
    $dll_to_register = $destination_path + $file

    if ( (file_exist $dll_to_register) )
    {
      & regsvr32 /u /s $dll_to_register  
      if(! $? )
      {
          throw "Failed to unregister $($dll_to_register)"
      }   
      else {
        Write-Host "Un-Registered $($dll_to_register)" 
      }
    }
      
    if (! (copy_file_to_win_sys $file)) 
    {
       return; 
    }
      
    & regsvr32 /s $dll_to_register 
    if(! $? )
    {
        throw "Failed to register $($dll_to_register)"
    } 

    Write-Host "Registered $($dll_to_register)"
}   

try
{
    $dll_name = $args[0]
    Write-Host $"Dll/Ocx/tlb: $($dll_name)" 

    $ret =copy_file_to_release ($dll_name + ".dll")
    $ret =copy_file_to_release ($dll_name + ".ocx")
    $ret =copy_file_to_release ($dll_name + ".oca")    
    $ret =copy_file_to_release ($dll_name + ".tlb")
    $ret =copy_file_to_release ($dll_name + ".lib")
    $ret =copy_file_to_release ($dll_name + ".dep")
    $ret =copy_file_to_release ($dll_name + ".dbg")
    $ret =copy_file_to_release ($dll_name + ".pdb")
    $ret =copy_file_to_release ($dll_name + ".exp")
        
    $ret =copy_file_to_win_sys_and_register ($dll_name + ".dll")
    $ret =copy_file_to_win_sys_and_register ($dll_name + ".ocx")
    $ret =copy_file_to_win_sys_and_register ($dll_name + ".tlb")

    $ret = copy_file_to_win_sys($dll_name + ".dbg")
    $ret =copy_file_to_win_sys($dll_name + ".pdb")


    Write-Host  "Finished";
}
catch 
{
    Write-Host "An error occurred:"
    Write-Host $_
}


