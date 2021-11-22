--Manager of the parent unit of parent unit of the recipient department
--Division manager two levels above the recipient

select hho.UID_PersonHead as UID_Person,
    dbo.QER_FGIPWORulerOrigin(hho.XObjectkey) as UID_PWORulerOrigin
    
from person p0
    join department d0 on p0.uid_department = d0.uid_department
    join department d1 on d1.uid_department = d0.uid_parentdepartment
    join department d2 on d2.uid_department = d1.uid_parentdepartment
-- also for delegations
	join HelperHeadOrg hho on hho.UID_Org = d2.uid_department
        
    join personWantsOrg pwo on p0.uid_person = pwo.uid_personOrdered
where pwo.UID_PersonWantsOrg = @UID_PersonWantsOrg