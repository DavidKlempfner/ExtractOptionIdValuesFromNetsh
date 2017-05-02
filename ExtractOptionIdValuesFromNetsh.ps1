<#
.SYNOPSIS
    This script helps extract option ID values from DHCP
 
.DESCRIPTION
    This script helps extract option ID values from DHCP
 
.INPUTS
 
.OUTPUTS
 
.EXAMPLE
    $netshOutputWithVendorClassAndBootpLast = '',
                    'Changed the current scope context to 11.13.32.0 scope.',
                    '',
                    'Options for Scope 11.13.32.0:',
                    '',
                    '      For vendor class [Microsoft Options]:',
                    '      OptionId : 2 ',
                    '      Option Value: ',
                    '             Number of Option Elements = 1',
                    '             Option Element Type = DWORD',
                    '             Option Element Value = 1',
                    '      OptionId : 3 ',
                    '      Option Value: ',
                    '             Number of Option Elements = 1',
                    '             Option Element Type = IPADDRESS',
                    '             Option Element Value = 11.13.32.1',
                    '      OptionId : 51 ',
                    '      Option Value: ',
                    '             Number of Option Elements = 1',
                    '             Option Element Type = DWORD',
                    '             Option Element Value = 31200',
                    '      OptionId : 60 ',
                    '      Option Value: ',
                    '             Number of Option Elements = 1',
                    '             Option Element Type = STRING',
                    '             Option Element Value = abc,
                    '      For user class [Default BOOTP Class]:',
                    '      OptionId : 51 ',
                    '      Option Value: ',
                    '             Number of Option Elements = 1',
                    '             Option Element Type = DWORD',
                    '             Option Element Value = 142800',
                    'Command completed successfully.'
    $optionId = '6'
 
    $values = GetValueFromOptionId $optionId $netshOutput
    #Outputs: '14.221.74.11', '14.221.76.34'
.NOTES
    Author: dklempfner@gmail.com
    Date: 23/01/2017
#>
 
function ExtractOptionIdValues
{
    Param([Parameter(Mandatory=$true)][Int32]$I,
          [Parameter(Mandatory=$true)][Object[]]$NetshScopeOptionsOutput)
 
    $values = New-Object 'System.Collections.Generic.List[String]'
 
    $numberOfRowsValueLineIsBelowOptionIdLine = 4
    $index = $I + $numberOfRowsValueLineIsBelowOptionIdLine
    $rowWithValue = $NetshScopeOptionsOutput[$index]
 
    while($rowWithValue.Contains('Option Element Value'))
    {
        $indexOfEqualsChar = $rowWithValue.IndexOf('=')
        $startIndexOfValue = $indexOfEqualsChar + 1
        $lengthOfValue = $rowWithValue.Length - $startIndexOfValue
        $values.Add($rowWithValue.Substring($startIndexOfValue, $lengthOfValue).Trim())
 
        $index++
        $rowWithValue = $NetshScopeOptionsOutput[$index]
    }
          
    return $values   
}
 
function GetStartIndexForDefaultBootpClass
{
    Param([Parameter(Mandatory=$true)][Object[]]$NetshScopeOptionsOutput)   
 
    for($i = 0; $i -lt $NetshScopeOptionsOutput.Count; $i++)
    {
        if($NetshScopeOptionsOutput[$i].Contains('For user class [Default BOOTP Class]'))
        {
            return $i
        }
    }
}
 
function GetValueFromOptionId
{
    Param([Parameter(Mandatory=$true)][String]$OptionId,
          [Parameter(Mandatory=$true)][Object[]]$NetshScopeOptionsOutput,
          [Parameter(Mandatory=$false)][Bool]$ShouldGetBootpValue = $false)
 
    $value = ''
 
    $indexOfBootp = GetStartIndexForDefaultBootpClass $NetshScopeOptionsOutput
    if(!$indexOfBootp)
    {
        $indexOfBootp = $NetshScopeOptionsOutput.Count - 1
    }
 
    if($OptionId -eq '51' -and $ShouldGetBootpValue)
    {       
        for($i = $indexOfBootp; $i -lt $NetshScopeOptionsOutput.Count; $i++)
        {
            if($NetshScopeOptionsOutput[$i].Contains("OptionId : $OptionId "))
            {           
                $value = ExtractOptionIdValues $i $NetshScopeOptionsOutput
            }                                              
        }
    }
    elseif($OptionId -eq '51')
    {       
        for($i = 0; $i -lt $indexOfBootp; $i++)
        {
            if($NetshScopeOptionsOutput[$i].Contains("OptionId : $OptionId "))
            {           
                $value = ExtractOptionIdValues $i $NetshScopeOptionsOutput
            }                                              
        }  
        if(!$value)
        {
            $numberOfTimesOption51WasFound = 0
            for($i = ($indexOfBootp + 1); $i -lt $NetshScopeOptionsOutput.Count; $i++)
            {
                if($NetshScopeOptionsOutput[$i].Contains('OptionId : 51 '))
                {   
                    $numberOfTimesOption51WasFound++       
                    if($numberOfTimesOption51WasFound -eq 2)
                    {
                        $value = ExtractOptionIdValues $i $NetshScopeOptionsOutput
                    }
                    
                }                                              
            }
        }
    }
    else
    {
        for($i = 0; $i -lt $NetshScopeOptionsOutput.Count; $i++)
        {
            #Space after $optionId is necessary to avoid getting option value 60 when option value 6, for example, is needed.
            if($NetshScopeOptionsOutput[$i].Contains("OptionId : $OptionId "))
            {           
                $value = ExtractOptionIdValues $i $NetshScopeOptionsOutput
            }                                              
        }
    }
    return $value
}