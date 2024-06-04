const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { CardFactory, MessageFactory } = require("botbuilder");

class Search {
  triggerPatterns = ["find", "search", "whereis", "locate"];

  async handleCommandReceived(context, message) {
    // verify the command arguments which are received from the client if needed.
    console.log(`App received message: ${message.text}`);

    // do something to process your command and return message activity as the response

    // render your adaptive card for reply message
    async function performSearch(queryStringg) {
      const searchEndpoint = "https://graph.microsoft.com/v1.0/search/query";
      const authToken =
        "eyJ0eXAiOiJKV1QiLCJub25jZSI6IlFUTlhDeUVGLVN4bUZqTDRDcUNpQ09sTWVoRlJYeFpFM1BxajR4N05Nc2siLCJhbGciOiJSUzI1NiIsIng1dCI6IkwxS2ZLRklfam5YYndXYzIyeFp4dzFzVUhIMCIsImtpZCI6IkwxS2ZLRklfam5YYndXYzIyeFp4dzFzVUhIMCJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8wMjFlZGI4Ni04NzU4LTRjNTYtOTM3Ny00ZTY2ZDIxOTZmYzAvIiwiaWF0IjoxNzE3NTMxMzQ1LCJuYmYiOjE3MTc1MzEzNDUsImV4cCI6MTcxNzYxODA0NSwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhXQUFBQVgzd2wvajNTeTdCUGFoVlFSUlBOMUd6R3psczA3VXlRVnRoaVcwdDZNNTlKZkpzWmk5UVpaSkNBMVRSeGxNUVZaZXR1RkJhbmRNcmZVQUY3STNDYk81WUFXbFhrS1MxRFVUbkNJZkRrclhrPSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggRXhwbG9yZXIiLCJhcHBpZCI6ImRlOGJjOGI1LWQ5ZjktNDhiMS1hOGFkLWI3NDhkYTcyNTA2NCIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiVW5ndXJlYW51IiwiZ2l2ZW5fbmFtZSI6IkFsZXgiLCJpZHR5cCI6InVzZXIiLCJpcGFkZHIiOiIxNDIuMTEyLjIzLjg3IiwibmFtZSI6IkFsZXggVW5ndXJlYW51Iiwib2lkIjoiNGI2MGU4MzQtNTJkOC00MDRjLWEyYzItZDkxNThjNDgxOWU4IiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDAzN0I4NDI4OTgiLCJyaCI6IjAuQVdJQWh0c2VBbGlIVmt5VGQwNW0waGx2d0FNQUFBQUFBQUFBd0FBQUFBQUFBQUNNQUFRLiIsInNjcCI6IkZpbGVzLlJlYWQuQWxsIG9wZW5pZCBwcm9maWxlIFVzZXIuUmVhZCBlbWFpbCBTaXRlcy5SZWFkLkFsbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IlBjT1pRY0hNbEJpTEJ1SHV4UEphSTFySllXeVNncU5ULXE0dVpXTmZRaWsiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiIwMjFlZGI4Ni04NzU4LTRjNTYtOTM3Ny00ZTY2ZDIxOTZmYzAiLCJ1bmlxdWVfbmFtZSI6ImFsZXhAYXVnc2lnbmFscy5jb20iLCJ1cG4iOiJhbGV4QGF1Z3NpZ25hbHMuY29tIiwidXRpIjoibFY2RjFLTTVmVXFyZTZHbmZoWHZBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19jYyI6WyJDUDEiXSwieG1zX3NzbSI6IjEiLCJ4bXNfc3QiOnsic3ViIjoiTWRhX1FjZ0Y4SkxwclZoWGE4RC1rNlJVZExZemhwa0V4SmZFamh4MFNwSSJ9LCJ4bXNfdGNkdCI6MTU2NDU4NTExNn0.KJE1JgZt9e0mee8mv72QbKSGiB_FCqg_-J9G3dABneX1AW_bH1egJCsrgqTa1n4AarWsR0nkwFK9cWc8zkpkRKqDhjGXacD5EWYyKwE-zn9WBQ4LLO8DQQ_UsIum9HCXpTrfytS1syb2WFzQNfxiv_yYconmvJJ_kAYbJa6pB-o_N14v53jzkOlOJtWMTKDwrAYUq4fVwIP6xEuYAYPRMGVM7FvGA2ibToNkQo6zfu3PsllWA6JiVjGEGPVmodQeV4i9kFyFcfiy8Nc10nezuEIJ-1wlGPoCgYQSkESflxpqlV80QJUh73s0wnVu3_17LJ_VPjumZVP-tvcAZo52mQ";

      const requestBody = {
        requests: [
          {
            entityTypes: ["driveItem"],
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

    const searchResults = await performSearch(message.text);

    let cardData;
    let searchURL;

    let allResults = [{
          type: "TextBlock",
          size: "Medium",
          weight: "Bolder",
          text: "${title}",
        },];

    let i = 0;

    await searchResults.value[0].hitsContainers[0].hits.forEach((result) => {
      i++;
      console.log(result);
      allResults.push({ 
      type: "TextBlock",
      text: `${i}: [${result.summary}](${result.resource.webUrl})`,
      size: "Small",
      wrap: true,
       },)
    });

    if (searchResults.value[0].hitsContainers[0].total === 0) {
      allResults = {
        title: "No results found",
      };

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
      version: "1.0",
      body: allResults,
      actions: [
        {
          type: "Action.OpenUrl",
          title: "Go to result",
          url: `${searchURL}`,
        },
        {
          type: "Action.OpenUrl",
          title: "Knowledge Base",
          url: "http://docs.aug.local:8888",
        },
      ],
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      version: "1.4",
    }).render(cardData);
    return MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
  }
}

module.exports = {
  Search,
};
