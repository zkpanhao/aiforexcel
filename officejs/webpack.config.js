const path = require("path");
const fs = require("node:fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");
const { DefinePlugin } = require("webpack");
const { handleCommonCommands } = require("./scripts/common-commands.cjs");
const { handleConnection } = require("./scripts/connection-settings.cjs");
const { clientLogRecord } = require("./scripts/local-server.cjs");
const commandRoot = path.join(process.env.USERPROFILE, "AIForExcel", "officejs");

module.exports = async (_env, argv) => ({
  entry: "./src/taskpane/taskpane.ts",
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "taskpane.js",
    clean: true,
  },
  resolve: { extensions: [".ts", ".js"] },
  module: {
    rules: [
      { test: /\.ts$/, use: "ts-loader", exclude: /node_modules/ },
      { test: /\.css$/, use: ["style-loader", "css-loader"] },
    ],
  },
  plugins: [
    new DefinePlugin({ __HOST_SMOKE_TEST__: JSON.stringify(argv.mode === "development" && process.env.AI_EXCEL_HOST_SMOKE_TEST === "1") }),
    new CopyWebpackPlugin({ patterns: [{ from: "assets", to: "assets" }, { from: "manifest.xml", to: "manifest.xml" }] }),
    new HtmlWebpackPlugin({
      template: "./src/taskpane/taskpane.html",
      filename: "taskpane.html",
      inject: "body",
    }),
    new HtmlWebpackPlugin({ template: "./src/commands/commands.html", filename: "commands.html", inject: false }),
  ],
  devServer: {
    port: 3000,
    server: { type: "https", options: argv.mode === "development" ? await devCerts.getHttpsServerOptions() : {} },
    host: "localhost",
    allowedHosts: ["localhost"],
    client: { overlay: { errors: true, warnings: false } },
    setupMiddlewares(middlewares, server) {
      server.app.use((request, response, next) => { if (!handleCommonCommands(commandRoot, request, response) && !handleConnection(commandRoot, request, response)) next(); });
      server.app.post("/__dev/office-test-result", (request, response) => {
        if (request.headers.origin !== "https://localhost:3000") return response.sendStatus(403);
        let body = "";
        request.on("data", chunk => { body += chunk; if (body.length > 3000) request.destroy(); });
        request.on("end", () => {
          try {
            const data = JSON.parse(body);
            const result = { at: new Date().toISOString(), passed: data.passed === true, detail: String(data.detail ?? "").slice(0, 2000) };
            fs.mkdirSync(path.join(__dirname, "work"), { recursive: true });
            fs.writeFileSync(path.join(__dirname, "work", "office-smoke.json"), JSON.stringify(result, null, 2));
            response.sendStatus(204);
          } catch { response.sendStatus(400); }
        });
      });
      server.app.post("/__dev/office-ready", (request, response) => {
        if (request.headers.origin !== "https://localhost:3000") return response.sendStatus(403);
        let body = "";
        request.on("data", chunk => { body += chunk; if (body.length > 1000) request.destroy(); });
        request.on("end", () => {
          try {
            const data = JSON.parse(body);
            const status = { at: new Date().toISOString(), host: data.host === "Excel" ? "Excel" : "Other", supported: data.supported === true, selectionReady: data.selectionReady === true };
            fs.mkdirSync(path.join(__dirname, "work"), { recursive: true });
            fs.writeFileSync(path.join(__dirname, "work", "office-ready.json"), JSON.stringify(status, null, 2));
            response.sendStatus(204);
          } catch { response.sendStatus(400); }
        });
      });
      server.app.post("/__dev/client-log", (request, response) => {
        if (request.headers.origin !== "https://localhost:3000") return response.sendStatus(403);
        let body = "";
        request.on("data", chunk => { body += chunk; if (body.length > 10000) request.destroy(); });
        request.on("end", () => {
          try {
            const data = JSON.parse(body);
            fs.mkdirSync(path.join(__dirname, "work"), { recursive: true });
            fs.appendFileSync(path.join(__dirname, "work", "client.log"), `${JSON.stringify(clientLogRecord({ release: "development" }, data))}\n`);
            response.sendStatus(204);
          } catch { response.sendStatus(400); }
        });
      });
      return middlewares;
    },
    static: { directory: path.join(__dirname, "dist") },
  },
});
