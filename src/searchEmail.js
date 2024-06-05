const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { CardFactory, MessageFactory } = require("botbuilder");

class SearchEmail {
  triggerPatterns = ["find email", "find -e"];

  async handleCommandReceived(context, message) {   
    // verify the command arguments which are received from the client if needed.
    console.log(`App received message: ${message.text}`);

    // do something to process your command and return message activity as the response

    // render your adaptive card for reply message
    async function performSearch(queryStringg) {
      const searchEndpoint = "https://graph.microsoft.com/v1.0/search/query";
      const authToken =
        "eyJ0eXAiOiJKV1QiLCJub25jZSI6ImhmQUNsQ3VHYVBNVlZCa0d4MWx0SmoxVi1Kd2NsWlUtaE1UT0hUVnF6RG8iLCJhbGciOiJSUzI1NiIsIng1dCI6IkwxS2ZLRklfam5YYndXYzIyeFp4dzFzVUhIMCIsImtpZCI6IkwxS2ZLRklfam5YYndXYzIyeFp4dzFzVUhIMCJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8wMjFlZGI4Ni04NzU4LTRjNTYtOTM3Ny00ZTY2ZDIxOTZmYzAvIiwiaWF0IjoxNzE3NTk0NjY4LCJuYmYiOjE3MTc1OTQ2NjgsImV4cCI6MTcxNzY4MTM2OCwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhXQUFBQVhZS1c4dGc4WlRSK3o4ejdEQnFXbmk2SHFKN1FzSzZFLytvL0hvSjdzc0d4Y2hDY3NSTjc1MVV0OFowNDJId1ZRVFpoNFdxY3M4a09MSFBvMmN5cDdjcHFXVFVTM0RERHh5WDZQWTEwMVYwPSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggRXhwbG9yZXIiLCJhcHBpZCI6ImRlOGJjOGI1LWQ5ZjktNDhiMS1hOGFkLWI3NDhkYTcyNTA2NCIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiVW5ndXJlYW51IiwiZ2l2ZW5fbmFtZSI6IkFsZXgiLCJpZHR5cCI6InVzZXIiLCJpcGFkZHIiOiIxNDIuMTEyLjIzLjg3IiwibmFtZSI6IkFsZXggVW5ndXJlYW51Iiwib2lkIjoiNGI2MGU4MzQtNTJkOC00MDRjLWEyYzItZDkxNThjNDgxOWU4IiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDAzN0I4NDI4OTgiLCJyaCI6IjAuQVdJQWh0c2VBbGlIVmt5VGQwNW0waGx2d0FNQUFBQUFBQUFBd0FBQUFBQUFBQUNNQUFRLiIsInNjcCI6IkNhbGVuZGFycy5SZWFkIEZpbGVzLlJlYWQuQWxsIE1haWwuUmVhZCBvcGVuaWQgcHJvZmlsZSBTaXRlcy5SZWFkLkFsbCBVc2VyLlJlYWQgZW1haWwgQ2hhdC5SZWFkIiwic2lnbmluX3N0YXRlIjpbImttc2kiXSwic3ViIjoiUGNPWlFjSE1sQmlMQnVIdXhQSmFJMXJKWVd5U2dxTlQtcTR1WldOZlFpayIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6IjAyMWVkYjg2LTg3NTgtNGM1Ni05Mzc3LTRlNjZkMjE5NmZjMCIsInVuaXF1ZV9uYW1lIjoiYWxleEBhdWdzaWduYWxzLmNvbSIsInVwbiI6ImFsZXhAYXVnc2lnbmFscy5jb20iLCJ1dGkiOiI1ODdFV2FrS1hFcWxYSzV2QmhaR0FBIiwidmVyIjoiMS4wIiwid2lkcyI6WyJiNzlmYmY0ZC0zZWY5LTQ2ODktODE0My03NmIxOTRlODU1MDkiXSwieG1zX2NjIjpbIkNQMSJdLCJ4bXNfc3NtIjoiMSIsInhtc19zdCI6eyJzdWIiOiJNZGFfUWNnRjhKTHByVmhYYThELWs2UlVkTFl6aHBrRXhKZkVqaHgwU3BJIn0sInhtc190Y2R0IjoxNTY0NTg1MTE2fQ.USAUqCYrKE6r0fNhhW5u0YnobdpxEM0Lhm1HVMvhsgiuDeWv1BV9OWshWA2_zDgJHnzkedWFEqRJk3DFWb2DoDaAOuKqR0_zlVIzgneoxJZnGj8v6TcCN3wFQWE33K0RzRZU1Sdb8iceRM2b6K3ajzB78gjAroOh76sxrjDMuoyzwsj1WSvQVnYQcoKGMRBesQWu_p5zuPBAS5yLVdVneQfB9DFcLAF6KSNOWPbXlUvsGGXZmiiCubJW9bhWBf1eGpWzq5jSbvuOWSV8i7F8PjJGYtEmvn_oanXm1y7beo__w-Y-enurk4-74DGgsBSBYU4jC9pfwdzj7ba0FpOtMg";

      const requestBody = {
        requests: [
          {
            entityTypes: ["message"],
            query: {
              queryString: queryStringg,
            },
          },
        ],
      };

      // Make a POST request to the search endpoint with the auth token
      const response = await fetch(searchEndpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${authToken}`,
        },
        body: JSON.stringify(requestBody),
      });

      // Process the response and return the result
      const result = await response.json();

      return result;
    }

    const searchResults = await performSearch(message.text.split(" ").slice(1).slice(1).join(" "));

    let cardData;
    let searchURL;

    let allResults = [
      {
        type: "TextBlock",
        size: "Medium",
        weight: "Bolder",
        text: "${title}",
      },
    ];

    let i = 0;

    console.log(searchResults);

    if (searchResults.value[0].hitsContainers[0].total != 0) {
      await searchResults.value[0].hitsContainers[0].hits.forEach((result) => {
        i++;
        console.log(result);
        allResults.push({
          type: "Container",
          items: [
            {
              type: "TextBlock",
              text: `${i}: ${result.summary}`,
              size: "Small",
              wrap: true,
            },
            {
              type: "ActionSet",
              actions: [
                {
                  type: "Action.OpenUrl",
                  title: "Open Link",
                  url: `${result.resource.webLink}`,
                },
              ],
            },
          ],
        });
      });
    }

    if (searchResults.value[0].hitsContainers[0].total === 0) {
      allResults = [
        {
          type: "TextBlock",
          size: "Medium",
          weight: "Bolder",
          text: "No results found.",
        },
        {
          type: "TextBlock",
          text: "Try searching for something else.",
        },
      ];

      searchURL = "http://docs.aug.local:8888";
    } else {
      cardData = {
        title: `Search: ${message.text.split(" ").slice(1).join(" ")}`,
      };

      searchURL = `${searchResults.value[0].hitsContainers[0].hits[0].resource.webUrl}`;
    }

    console.log(allResults);

    const cardJson = AdaptiveCards.declare({
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.3",
      body: allResults,
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.4",
    }).render(cardData);
    return MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
  }
}

module.exports = {
  SearchEmail,
};
