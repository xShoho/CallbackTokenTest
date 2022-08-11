/* eslint-disable prettier/prettier */
<reference path="../../frameworks/Scripts/office/1/office.js" />
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  Office.context.mailbox.getCallbackTokenAsync((result) => {
    const email = document.querySelector(".email");
    const username = document.querySelector(".username");
    const tokenres = document.querySelector(".token");

    if(result.status == Office.AsyncResultStatus.Failed) {
      email.innerHTML = "<b>Can't get token</b>";
      return;
    }
    let user = Office.context.mailbox.userProfile;

    email.innerHTML = `email: ${ user.emailAddress }`;
    username.innerHTML = `email: ${ user.displayName }`;

    tokenres.innerHTML = result.value;
  });
}
