## Microsoft Graph API Example

A SPFX Webpart for creating events in outlook from a Sharepoint list.

### Building the code

```bash
git clone the repo
npm install
npm install -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean
gulp test
gulp serve
gulp bundle
gulp package-solution
