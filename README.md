## aadgroup-webpart

A Webpart for Sharepoint online similar to the standard People-Webpart.
In edit-Mode the user can choose a group from the Azure Active Directory.
The Webpart then shows all the members of this group.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
