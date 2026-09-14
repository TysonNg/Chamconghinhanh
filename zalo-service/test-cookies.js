const fs = require("fs");
const creds = JSON.parse(fs.readFileSync("./credentials.json", "utf-8"));
console.log(creds.cookie.map(c => ({ key: c.key || c.name, domain: c.domain })));
