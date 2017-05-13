/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

// This sample uses an open source OAuth 2.0 library that is compatible with the Azure AD v2.0 endpoint. 
// Microsoft does not provide fixes or direct support for this library. 
// Refer to the library’s repository to file issues or for other support. 
// For more information about auth libraries see: https://azure.microsoft.com/documentation/articles/active-directory-v2-libraries/ 
// Library repo: https://github.com/MrSwitch/hello.js

"use strict";

(function () {
    angular
      .module('app')
      .service('GraphHelper', ['$http', function ($http) {

          // Initialize the auth request.
          hello.init({
              aad: clientId // from public/scripts/config.js
          }, {
              redirect_uri: redirectUrl,
              scope: graphScopes
          });

          return {

              // Sign in and sign out the user.
              login: function login() {
                  hello('aad').login({
                      display: 'page',
                      state: 'abcd'
                  });
              },
              logout: function logout() {
                  hello('aad').logout();
                  delete localStorage.auth;
                  delete localStorage.user;
              },
              close: function close() {
                  window.close();
              },

              // Get the profile of the current user.
              me: function me() {
                  return $http.get('https://graph.microsoft.com/v1.0/me');
              },

              // Send an email on behalf of the current user.
              sendMail: function sendMail(email) {
                  return $http.post('https://graph.microsoft.com/v1.0/me/sendMail', { 'message': email, 'saveToSentItems': true });
              },

              // Upload a file on behalf of the current user.
              uploadFile: function uploadFile(fileName, fileContent) {
                  var Promise = Promise || ES6Promise.Promise; // do this to access Promise directly
                  return new Promise(function (resolve, reject) {
                      var url = "https://graph.microsoft.com/v1.0/me/drive/root:/Documents/" + fileName + ":/content";
                      let auth = angular.fromJson(localStorage.auth);
                      var oReq = new XMLHttpRequest();
                      oReq.open("PUT", url, true);
                      oReq.setRequestHeader("Content-Type", "text/plain");
                      oReq.setRequestHeader("Authorization", "Bearer " + auth.access_token);
                      oReq.onload = function () {
                          if (this.status >= 200 && this.status < 300) {
                              resolve(oReq.response);
                          } else {
                              reject({
                                  status: this.status,
                                  statusText: oReq.statusText
                              });
                          }
                      };
                      oReq.onerror = function () {
                          reject({
                              status: this.status,
                              statusText: oReq.statusText
                          });
                      };

                      oReq.send(fileContent);
                  });
              },

              // Send a graph request on behalf of the current user.
              exploreGraph: function exploreGraph(graphUrl, graphBody) {
                  return $http.post(graphUrl, graphBody);
              }
          }
      }]);
})();