﻿<#

Approval URL:



https://login.windows.net/common/oauth2/authorize?response_type=code&resource=https%3A%2F%2Fmanage.office.com&client_id=e4b98c10-fbff-4e1f-9b60-f3936edd2e9a&redirect_uri=https%3A%2F%2Fhappapp1.mgroup1.phydeauxnet.net

Approval URL response

https://happapp1.mgroup1.phydeauxnet.net/?code=AQABAAIAAABHh4kmS_aKT5XrjzxRAtHzlpGgV_KVOw4EqGtl1rwbpOYg_Cv9zRP69KPouYoxrzAqOvF152CxftGYX7Jtz-48B24cxdcdN1IfyDLOhuQZU3aVUHAldPkbYDXzKNpfpvZCGPQXXjNN7Pi9SrMLvgVdumrXKHQOk5PNLEByQbO6IR8aBHtoIHUtqVSZraHJD_5oU5NvQQQH1R9SXhxWfbZ8sw9AfNQLK52DH7zjUWTDBkEqwhxk2Wb1XanII0aTp3fTrOCMsFlHSfe4Qzk4BQ_K1jl7blCvbZ0Tyz4uKRIele46UrPNC_tzYRz_09_jR0SmgRjtopA1kquw2Lk44otu8zRckCYw2GRswytDJbU3Ves0cd93n9JnUL0AqlDo4uM6-LvOGuCNu-7t5kZqNOYlzvnH_Qubm9KntwMXAxJzbZWkmbC2Hj8UCd5JnbdI7LKCqf4FEE0tOok9SK59fZNwAMEAM9e8Ag4kZ_K3HRxB3czruXkcjIhHx3uGfnH69hbhSoN9EO_TuP_bbDVr4eLI-FqQlUSuLKlzN9EtcoXlFCJLan61ihp9_GT0zROI1nsgAA&session_state=a0ce0439-1fc1-4db8-b731-935458d4f0ae

#>


$ClientID = 'e4b98c10-fbff-4e1f-9b60-f3936edd2e9a'

$ClientKey = 'koMeV4vGziC6yO9DbeMqtFe2I/9k9dzZTVrJb11HBd0='

$AuthCode = 'AQABAAIAAABHh4kmS_aKT5XrjzxRAtHzJDDM4Uk4qR_7ZmRZDX7l0yq_udCoDEAHLKgmI40jUuUvbZKD0jjRLqPROUAWEWujl4HaSFih8mKcik18VHODSL6VfGi6KyhSo7po19dfpVTzrDtqYYMxy2HQBWFbqjTYvkI47Lh9SKCkw2JgY0df7asMy9r_rMqXkTMPpMbg7HpbGuYJu4xhWrisRXm97VC03VUVL3pxX5VZiLek5n3lI0A1mjvmbeAQSo4Kgjde4NghH-DOFnjG7ZgBOzmWdHwLRB50bwfdrOJCTJDtrKVxbwdjo7WFIfOBtn-aF-tWjqhu48ayU6C_p0o0KpyY_eireIUanHwcX6C6XmzTT11JOyAs10q-_4Pe6PFImmdpdInnDIz4-CenSi7ffXJndZpFlFWOH1WncHEDdgJmx61WGZ6VDJQKcR_Vmt5P0cZfePs5mTpD58MwWn6nuiwtgCHA4RHDMyN4mzEi4sl8dmyADTjIIinXVy7mnxGra-r-RTj0f7abL62zyjdeq1RRVrZMi9IgwVpZ8OiOuEScrX1TE3JOBgOi7L9allkIWFVuNOYmxUZ2fwcS6oms_wvQ58rkIAA&session_state=8620a708-3417-49cf-8309-58f0f504205e'

$PostBody = "resource=https%3A%2F%2Fmanage.office.com&client_id=$ClientID&redirect_uri=https%3A%2F%2Fhappapp1.mgroup1.phydeauxnet.net&client_secret=$ClientKey&grant_type=authorization_code&code=$AuthCode"

$Token = Invoke-WebRequest -Uri https://login.windows.net/common/oauth2/token -Method Post -Body $PostBody
