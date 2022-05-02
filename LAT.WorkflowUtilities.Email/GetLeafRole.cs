using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

//This activity is designed to return the leaf role id
//given the input Business Unit and Security Role.

namespace LAT.WorkflowUtilities.Email
{
    public class GetLeafRole : WorkFlowActivityBase
    {
        public GetLeafRole() : base(typeof(GetLeafRole)) { }

        [RequiredArgument]
        [Input("Business Unit")]
        [ReferenceTarget("businessunit")]
        public InArgument<EntityReference> BusinessUnit { get; set; }

        [RequiredArgument]
        [Input("Security Role")]
        [ReferenceTarget("role")]
        public InArgument<EntityReference> SecurityRole { get; set; }

        [Output("Role String")]
        public OutArgument<String> OutRoleString { get; set; }

        [Output("Role Entity Reference")]
        [ReferenceTarget("role")]
        public OutArgument<EntityReference> OutRoleReference { get; set; }


        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));
            if (SecurityRole == null)
                throw new ArgumentNullException("Recipient Role is null.");
            if (BusinessUnit == null)
                throw new ArgumentNullException("Business Unit is null.");

            //Get the input parameter values as entity reference
            EntityReference InputRole = SecurityRole.Get(context);
            EntityReference InputBusinessUnit = BusinessUnit.Get(context);

            if (InputRole == null || InputRole.Id == Guid.Empty)
                throw new InvalidWorkflowException("Invalid Role GUID");
            if (InputBusinessUnit == null || InputBusinessUnit.Id == Guid.Empty)
                throw new InvalidWorkflowException("Invalid Business GUID");

            //Get the parent root role id and business unit id from input values
            EntityReference pRoleID = (EntityReference)localContext.OrganizationService.Retrieve("role", InputRole.Id, new ColumnSet("parentrootroleid")).Attributes["parentrootroleid"];
            Guid ParentRootRole = pRoleID.Id;
            Guid InBusinessUnitID = InputBusinessUnit.Id;

            //get the correct role id for the business unit
            QueryExpression query = new QueryExpression("role");
            query.ColumnSet = new ColumnSet("roleid");
            query.Criteria.AddCondition("businessunitid", ConditionOperator.Equal, InBusinessUnitID);
            query.Criteria.AddCondition("parentrootroleid", ConditionOperator.Equal, ParentRootRole);
            EntityCollection BuRoles = localContext.OrganizationService.RetrieveMultiple(query);

            if (BuRoles != null && BuRoles.Entities.Count > 0 && BuRoles.Entities[0].Contains("roleid"))
            {
                EntityReference resultReference = localContext.OrganizationService.Retrieve("role", BuRoles.Entities[0].GetAttributeValue<Guid>("roleid"), new ColumnSet("roleid")).ToEntityReference();
                var @resultString = BuRoles.Entities[0].GetAttributeValue<Guid>("roleid").ToString();
                OutRoleString.Set(context, @resultString);
                OutRoleReference.Set(context, resultReference);
            }
            else
            {
                OutRoleString.Set(context, null);
                OutRoleReference.Set(context, null);
            }

        }
    }
}

