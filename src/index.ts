import { Client, getGraphRestSDKClient } from "@microsoft/microsoft-graph-client";

/** Imports for authentication */
import { ClientSecretCredential } from "@azure/identity";
import { AzureIdentityAuthenticationProvider } from "@microsoft/kiota-authentication-azure";
import { tenantId, clientId, clientSecret, scopes } from "./secrets";

import "@microsoft/microsoft-graph-client/users"
import "@microsoft/microsoft-graph-client/groups/group"
import { MicrosoftGraphGroup } from "@microsoft/microsoft-graph-types";
const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);


const authProvider = new AzureIdentityAuthenticationProvider(credential, [scopes]);

const client = Client.init({
    authProvider,
});


async function nonTypedRequest(){
    const response = await client.api("/users").get();
   // console.log(response);
}

/** Create the typedClient */
export const typedClient = getGraphRestSDKClient(client);
async function getUsers(){
    const users = await typedClient.api("/users").get();
    users.value.forEach((user)=> {console.log(user.givenName)})
}

async function patchUser(){
    const patchResponse = await typedClient.api("/users/{user-id}","02b562bf-51c3-4a2e-897f-085a0c3f8259").patch({
        "department": "Sales & Marketing"
    })

    console.log(patchResponse)
}

async function createGroup(){
    const group: MicrosoftGraphGroup = {
        "displayName": "Library Assist",
        "mailEnabled": true,
        "mailNickname": "demoExample",
        "securityEnabled": true,
        "groupTypes": [
            "Unified"
        ]
    }
    const postGroup = await typedClient.api("/groups").post(group);
    console.log(postGroup);
}

// nonTypedRequest().then();
// patchUser().then();
// getUsers().then();
createGroup().then();
