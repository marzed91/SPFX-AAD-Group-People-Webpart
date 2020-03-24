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

### Todo

- [ ] Make title optional:
By default, the title should be the Groupname, but users should be able to change or deactivate it.

- [ ] Handle nested groups:
As of now, users in nested groups are ignored and the group is shown like a user but without any details.
All users should be shown independantly of wheter they are in a nested group or not. Nested groups however don't need to be shown.