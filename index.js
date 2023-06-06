const express = require("express");
const request = require( "request-promise");
const dotenv = require("dotenv");
dotenv.config();
const app = express();
const port = process.env.PORT || 3000;
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
class RedTech {
  static TenantName = "redtechly0";
  static ClientID = "6849427b-5ac7-4342-b067-77b621f8f5ff";
  static ClientSecret = "hZVtEl2dB6QJ1bvCCqXZzbwL2YtR9dKHP26zfGPm3yM=";
  static TenantID = "e1cbaea0-f8c8-4ce7-a484-ade672f55f8d";
  static ApplicationID = "00000003-0000-0ff1-ce00-000000000000";
  static RefreshToken =
    "PAQABAAEAAAD--DLA3VO7QrddgJg7WevrouvTWRodkELn8lccr27sIuWXUsmlozjz3T6b0kD885RwG08CjphqnKCc-2eEYGoCLoUcdUpYZ82pvEAo-WZTMzW7eoxrV6KC1j5nuhcx7lIxsOK1QrfZp3AJol6LReGnWUt8Wc7BYfva2vlBz3MB4Y3QoHR8AmN3z4V4vQ0YGiwUklWEFtP9RhMMaMxCXM250dYq23ptyF4_MKEgadp0QtjxZxO82GEd8dbtsp8ED1-4fmpKfxIFoH_-QRQIzHzSBN0s2PmSCE1C2o7ezoXFvyAA";
  static AccessToken;
  static IssueDate;
  constructor(ListName, SiteName) {
    this.SiteName = SiteName;
    this.ListName = ListName;
  }
  static async generatetoken() {
    try {
      let result = await request({
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        form: {
          grant_type: "refresh_token",
          client_id: this.ClientID + "@" + this.TenantID,
          client_secret: this.ClientSecret,
          resource:
            this.ApplicationID +
            "/" +
            this.TenantName +
            ".sharepoint.com@" +
            this.TenantID,
          refresh_token: this.RefreshToken,
        },
        uri:
          "https://accounts.accesscontrol.windows.net/" +
          this.TenantID +
          "/tokens/OAuth/2",
        method: "POST",
      });
      return JSON.parse(result).access_token;
    } catch (err) {
      console.log(err);
    }
  }
  static async getAccessToken() {
    if (
      this.AccessToken == null ||
      Date.now().valueOf() > this.IssueDate + 28500000
    ) {
      console.log("Generating new token");
      this.IssueDate = Date.now().valueOf();
      this.AccessToken = await this.generatetoken();
    }
    return this.AccessToken;
  }
  async getListItems(Field, Value) {
    try {
      let result = await request({
        json: true,
        verbose: true,
        headers: {
          Authorization: "Bearer " + (await RedTech.getAccessToken()),
          "Content-Type": "application/json; odata=verbose",
          Accept: "application/json; odata=nometadata",
        },
        uri: encodeURI(
          "https://" +
            RedTech.TenantName +
            ".sharepoint.com/sites/" +
            this.SiteName +
            "/_api/web/lists/GetByTitle('" +
            this.ListName +
            "')/items" +
            "?$filter= " +
            Field +
            " eq '" +
            Value +
            "'"
        ),
        body: "",
        method: "GET",
      });
      return result;
    } catch (err) {
      console.log(err);
    }
  }
  async getListAllItems() {
    try {
      let result = await request({
        json: true,
        verbose: true,
        headers: {
          Authorization: "Bearer " + (await RedTech.getAccessToken()),
          "Content-Type": "application/json; odata=verbose",
          Accept: "application/json; odata=nometadata",
        },
        uri: encodeURI(
          "https://" +
            RedTech.TenantName +
            ".sharepoint.com/sites/" +
            this.SiteName +
            "/_api/web/lists/GetByTitle('" +
            this.ListName +
            "')/items"
        ),
        body: "",
        method: "GET",
      });
      return result;
    } catch (err) {
      console.log(err);
    }
  }
}
app.get("/api" ,async (req, res) => {
  const Value = req.body.key;
  const Field = "field_1";
  const SiteName = "REDRestaurants";
  const ListName = "SaqidaniCart";
  let Saqidani = new RedTech(ListName, SiteName);
  const lists = (await Saqidani.getListItems(Field, Value)).value;
  const Products = lists.map((list) => {
    return {
      title: `${list.field_2}  العدد  ${list.field_4}`,
      subtitle: list.field_3,
      image_url: list.field_7,
      action_url: "https://manychat.com",
      buttons: [
        {
          type: "node",
          caption: "حدف",
          target: "Delete",
          actions: [
            {
              action: "set_field_value",
              field_name: "ProductID",
              value: list.field_5,
            },
            {
              action: "set_field_value",
              field_name: "ProductName",
              value: list.field_2,
            },
            {
              action: "set_field_value",
              field_name: "ProductPrice",
              value: list.field_6,
            },
          ],
        },
      ],
    };
  });
  const total = lists.reduce((acc, list) => {
    return acc + list.field_6 * list.field_4;
  }, 0);
  console.log(total)
  const delivery = 10;
  const response = {
    version: "v2",
    content: {
      messages: [
        {
          type: "cards",
          elements: Products,
          image_aspect_ratio: "square",
        },
        {
          type: "text",
          text: `العناصر في السلة 🛒 : ${total} دينار  \n التوصيل 🛵 : ${delivery} دينار \n ---------------------------- \n المجموع : ${
            total + delivery
          } دينار`,
          buttons: [
            {
              type: "node",
              caption: "إتمام عملية الشراء ✅",
              target: "CollectData",
            },
            {
              type: "node",
              caption: "إضافة عنصر ➕",
              target: "Back",
            },
            {
              type: "node",
              caption: "حذف الكل ❌",
              target: "DeleteAll",
            },
          ],
        },
      ],
      actions: [
        {
          action: "set_field_value",
          field_name: "Total",
          value: `${total+delivery}`,
        },
      ],
      quick_replies: [],
    },
  };
  res.json(response);
});
app.get("/api/all", async (req, res) => {
  const ListName = "SaqidaniCart";
  const SiteName = "REDRestaurants";
  let Saqidani = new RedTech(ListName, SiteName);
  res.json(await Saqidani.getListAllItems());
});
app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`);
});
