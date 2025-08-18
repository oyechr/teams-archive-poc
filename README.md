
# Local Development

To run the app locally, follow these steps:

1. **Install dependencies**

   - Run:
     ```
     npm install
     ```
2. **Set up your .env file**

   - Create a `.env` file in the project root and add your own variables. Example:
     ```env
     # The public domain name of where you host your application
     PUBLIC_HOSTNAME=<your-ngrok-domain>

     # Id of the Microsoft Teams application
     APPLICATION_ID=<your-app-id>
     AZURE_AD_CLIENT_SECRET=<your-client-secret>
     AZURE_AD_TENANT_ID=<your-tenant-id>
     # Package name of the Microsoft Teams application
     PACKAGE_NAME=teamsarchivepoc

     # App Id and App Password for the Bot Framework bot
     MICROSOFT_APP_ID=
     MICROSOFT_APP_PASSWORD=

     # Port for local debugging
     PORT=3007

     # Security token for the default outgoing webhook
     SECURITY_TOKEN=

     # ID of the Outlook Connector
     CONNECTOR_ID=

     # Application Insights instrumentation key
     APPINSIGHTS_INSTRUMENTATIONKEY=

     # NGROK configuration for development
     # NGROK authentication token (leave empty for anonymous)
     NGROK_AUTH=
     # NGROK sub domain. ex "myapp" or  (leave empty for random)
     NGROK_SUBDOMAIN=
     # NGROK region. (us, eu, au, ap - default is us)
     NGROK_REGION=

     # Debug settings, default logging "msteams"
     DEBUG=msteams

     TAB_APP_ID=<your-app-id>
     TAB_APP_URI=api://<your-ngrok-domain>/<your-app-id>
     ```

3. **Upload the manifest to Teams**

- Run:
  ```
  gulp manifest --no-schema-validation
  ```
- This will generate the Teams app manifest package for uploading to Teams.
4. **Start ngrok tunnel**

   - Replace `<your-ngrok-domain>` with your own ngrok domain.
   - Example:
     ```
     ngrok http --url=<your-ngrok-domain> 3007
     ```
     
5. **In another terminal serve the app locally with gulp**

   - Run:
     ```
     gulp serve --debug
     ```

This will start the local server and expose it via ngrok for Teams integration and testing.

6. **Open the app in your Teams client**
  - In Microsoft Teams, go to 'Apps' and select your uploaded app to start testing.

# teams archive poc - Microsoft Teams App

Generate a Microsoft Teams application.

TODO: Add your documentation here

## Getting started with Microsoft Teams Apps development

Head on over to [Microsoft Teams official documentation](https://developer.microsoft.com/en-us/microsoft-teams) to learn how to build Microsoft Teams Tabs or the [Microsoft Teams Yeoman generator docs](https://github.com/PnP/generator-teams/docs) for details on how this solution is set up.

## Project setup

All required source code are located in the `./src` folder:

* `client` client side code
* `server` server side code
* `public` static files for the web site
* `manifest` for the Microsoft Teams app manifest

For further details see the [Yo Teams documentation](https://github.com/PnP/generator-teams/docs)

## Building the app

The application is built using the `build` Gulp task.

```bash
npm i -g gulp-cli
gulp build
```

## Building the manifest

To create the Microsoft Teams Apps manifest, run the `manifest` Gulp task. This will generate and validate the package and finally create the package (a zip file) in the `package` folder. The manifest will be validated against the schema and dynamically populated with values from the `.env` file.

```bash
gulp manifest
```

## Deploying the manifest

Using the `yoteams-deploy` plugin, automatically added to the project, deployment of the manifest to the Teams App store can be done manually using `gulp tenant:deploy` or by passing the `--publish` flag to any of the `serve` tasks.

## Configuration

Configuration is stored in the `.env` file.

## Debug and test locally

To debug and test the solution locally you use the `serve` Gulp task. This will first build the app and then start a local web server on port 3007, where you can test your Tabs, Bots or other extensions. Also this command will rebuild the App if you change any file in the `/src` directory.

```bash
gulp serve
```

To debug the code you can append the argument `debug` to the `serve` command as follows. This allows you to step through your code using your preferred code editor.

```bash
gulp serve --debug
```

## Useful links

* [Debugging with Visual Studio Code](https://github.com/pnp/generator-teams/blob/master/docs/docs/user-guide/vscode.md)
* [Developing with ngrok](https://github.com/pnp/generator-teams/blob/master/docs/docs/concepts/ngrok.md)
* [Developing with Github Codespaces](https://github.com/pnp/generator-teams/blob/master/docs/docs/user-guide/codespaces.md)

## Additional build options

You can use the following flags for the `serve`, `ngrok-serve` and build commands:

* `--no-linting` or `-l` - skips the linting of Typescript during build to improve build times
* `--debug` - builds in debug mode and significantly improves build time with support for hot reloading of client side components
* `--env <filename>.env` - use an alternate set of environment files
* `--publish` - automatically publish the application to the Teams App store

## Deployment

The solution can be deployed to Azure using any deployment method.

* For Azure Devops see [How to deploy a Yo Teams generated project to Azure through Azure DevOps](https://www.wictorwilen.se/blog/deploying-yo-teams-and-node-apps/)
* For Docker containers, see the included `Dockerfile`

## Logging

To enable logging for the solution you need to add `msteams` to the `DEBUG` environment variable. See the [debug package](https://www.npmjs.com/package/debug) for more information. By default this setting is turned on in the `.env` file.

Example for Windows command line:

```bash
SET DEBUG=msteams
```

If you are using Microsoft Azure to host your Microsoft Teams app, then you can add `DEBUG` as an Application Setting with the value of `msteams`.
