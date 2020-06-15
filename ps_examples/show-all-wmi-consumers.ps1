Get-WMIObject -Namespace root\Subscription -Class __EventFilter | where {$_.Name -notlike "SCM*"}
Get-WMIObject -Namespace root\Subscription -Class __FilterToConsumerBinding  | where {$_.Name -notlike "SCM*"}
Get-WMIObject -Namespace root\Subscription -Class __EventConsumer  | where {$_.Name -notlike "SCM*"}