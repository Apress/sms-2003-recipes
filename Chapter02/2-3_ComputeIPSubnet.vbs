'adopted from the SMS 2003 Scenarios and Procedures guide
strIPAddress = inputbox("Enter IP Address")
strSubnetMask = inputbox("Enter Subnet Mask")

dim addressbytes(4)
dim subnetmaskbytes(4)
i=0
period = 1
while period<>len( strIPAddress ) + 2
        prevperiod=period
        period = instr( period+1, strIPAddress, "." ) + 1
        if period = 1 then period = len( strIPAddress ) + 2
         addressbyte = _
             mid( strIPAddress, prevperiod, period-prevperiod-1 )
         addressbytes(i)=addressbyte
        i=i+1
wend

i=0
period = 1
while period<>len( strSubnetMask ) + 2
        prevperiod=period
        period = instr( period+1, strSubnetMask, "." ) + 1
        if period = 1 then period = len( strSubnetMask ) + 2
         subnetmaskbyte = _
             mid( strSubnetMask, prevperiod, period-prevperiod-1 )
         subnetmaskbytes(i)=subnetmaskbyte
        i=i+1
wend
for i=0 to 3
        subnet = subnet & _
            (addressbytes(i) AND subnetmaskbytes(i)) & "."
next
subnet = left( subnet, len(subnet)-1 )
msgbox "Subnet: " & subnet
