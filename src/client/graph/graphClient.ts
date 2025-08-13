// import { Client } from "@microsoft/microsoft-graph-client";
// import { authentication } from "@microsoft/teams-js";
// import axios from "axios";
// import "isomorphic-fetch";


// export async function getGraphTokenFromServer(): Promise<string> {

//   const ssoToken = await authentication.getAuthToken({
//     resources: [process.env.TAB_APP_URI as string],
//     silent: false,
//   });

//   const response = await axios.post("/api/getGraphToken", { ssoToken });
//   return response.data.access_token;
// }

// export async function fetchUserChats() {
//   const graphToken = await getGraphTokenFromServer();

//   const client = Client.init({
//     authProvider: (done) => {
//       done(null, graphToken);
//     },
//   });

//   return await client.api("/me/chats").get();
// }
