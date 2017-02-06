// Add your app's registration data here
var client_id = '';
var client_secret = '';


var express = require('express');
var session = require('express-session')
var request = require('request');
var https = require('https');
var http = require('http');

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
            content += "<pre>" + auth_code + "</pre>";
            content += '<a href="/refresh?code=' + result.refresh_token + '" target="_blank">Refresh Token</a>';

            requestSubscribe(result.access_token, function(subBody) {

                content += "</p>";
                content += "<pre>" + subBody + "</pre>";
                res.end(content)
            });
        })
    } else {
        // This is an OAUTH Implicit Grant workflow

        var token = {
            access_token: req.query.access_token,
            token_type: req.query.token_type,
            expires_in: req.query.expires_in,
            scope: req.query.scope
        }

        var content = "<pre>" + JSON.stringify(token, null, 2) + "</pre>"
        content += "<pre>" + token.access_token + "</pre>";
        content += '<a href="/refresh?code=' + result.refresh_token + '" target="_blank">Refresh Token</a>'
        res.end(content)

    }
});

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


function requestSubscribe(token, callback) {

    var expDate = new Date();
    expDate.setDate(expDate.getDate() + 1);


    var subscription_request = {
        url: 'https://graph.microsoft.com/v1.0/subscriptions/',
        headers: {
            Authorization: 'Bearer ' + token
        },
        json: {
            changeType: 'created',
            notificationUrl: 'https://localhost:3001/notificationClient',
            resource: 'me/messages',
            expirationDateTime: expDate.toISOString()
        }
    }

    request.post(subscription_request, function(err, httpResponse, body) {
        callback(body);
    });

}

app.get('/notificationClient', function(req, res) {
    res.send('Success');

});

app.post('/notificationClient', function(req, res) {

    if (req.query.validationToken != null) {
        res.set('Content-Type', 'text/plain');
        res.status(200).send(req.query.validationToken);
    } else {
        res.status(400).end();
    }

});


const fs = require('fs');

const options = {
    pfx: fs.readFileSync('localhost-personal.pfx'),
    passphrase: 'January5191'
};

http.createServer(app).listen(3000);
https.createServer(options, app).listen(3001);

// Start listening on port 3000
//app.listen(3000, function () {
console.log('Example app listening on port 3000!');
//});