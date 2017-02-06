// Add your app's registration data here
const client_id = '';
const client_secret = '';

const graph = require("@microsoft/microsoft-graph-client");
const express = require('express');
const session = require('express-session')
const request = require('request');
const https = require('https');
const http = require('http');

var app = express();
app.use('/static', express.static('public'));
app.use(session({
    secret: 'keyboard cat',
    cookie: {}
}))

// Define the various URIs we need
var auth_uri = ' https://login.microsoftonline.com/common/oauth2/authorize';
var token_uri = 'https://login.microsoftonline.com/common/oauth2/token';
var redirect_uri = 'http://localhost:3000/returned';

// Define scopes
// NOTE: You must request offline_access in order to recieve a refresh_token. Without it the
// the autorization will only live for a limited time (typically 1 hour).
var client_scopes = 'https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.Read offline_access';


// Forms
// There are two forms we post to the v2 endpoint, one to request the initial token and another 
// to request a new token using a previous refresh_token
var token_request = {
    form: {
        grant_type: 'authorization_code',
        code: '', // We set this at runtime
        client_id: client_id,
        client_secret: client_secret,
        redirect_uri: redirect_uri,
        resource: 'https://graph.microsoft.com/'
    }
}

var refresh_token_request = {
    form: {
        grant_type: 'refresh_token',
        refresh_token: '', // We set this at runtime
        client_id: client_id,
        client_secret: client_secret,
        redirect_uri: redirect_uri,
        resource: 'https://graph.microsoft.com/'
    }
}

// This is the web root and provides a link that kicks off the OAUTH process 
app.get('/', function(req, res) {
    var codegrant_endpoint = auth_uri + '?client_id=' + client_id + '&response_type=code&redirect_uri=' + redirect_uri;
    res.send('<div><a href="' + codegrant_endpoint + '" target="_blank">Code Grant Workflow</a></div>');
});


// This is the return point after we have executed the first OAUTH step.
// Here we convert the auth code returned by the endpoint into a bearer token 
// we can use to call the API. We also return a link to refresh this token 
// when it expires.
app.get('/returned', function(req, res) {

    if (req.query.code != null) {
        // This is an OAUTH Code Grant workflow

        // Grab the auth code from the query params
        var auth_code = req.query.code;

        // Add this auto_code to our token_request
        token_request.form.code = auth_code;

        // Post this form to the v2 Endpoint and display the result in the browser    
        request.post(token_uri, token_request, function(err, httpResponse, body) {
            var result = JSON.parse(body);

            var content = "<pre>" + JSON.stringify(result, null, 2) + "</pre>"
            content += "</p>";

            getProfile(result.access_token, function(profile) {
                content += "<pre>" + JSON.stringify(profile, null, 2) + "</pre>";
                content += "</p>";

                content += '<a href="/refresh?code=' + result.refresh_token + '" target="_blank">Refresh Token</a>';
                res.end(content)
            });
        });
    }
});

function getProfile(access_token, callback) {
    var client = graph.Client.init({
        authProvider: (done) => {
            done(null, access_token); //first parameter takes an error if you can't get an access token
        }
    });
    client
        .api('/me')
        .get((err, res) => {
            callback(res);
        });
}


// This is where we convert the refresh token into a usable 
// brearer token we can use for API calls. 
app.get('/refresh', function(req, res) {

    // Grab the refresh_token from the query params
    var refresh_token = req.query.code;

    // Add this auto_code to our refresh_token_request
    refresh_token_request.form.refresh_token = refresh_token;

    // Post this form to the v2 Endpoint and display the result in the browser
    request.post(token_uri, refresh_token_request, function(err, httpResponse, body) {
        var result = JSON.parse(body);

        var content = "<pre>" + JSON.stringify(result, null, 2) + "</pre>"
        content += "<pre>" + refresh_token + "</pre>";
        content += '<a href="/refresh?code=' + result.refresh_token + '" target="_blank">Refresh Token</a>'
        res.end(content)
    })
});


http.createServer(app).listen(3000);

console.log('Example app listening on port 3000!');