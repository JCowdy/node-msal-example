const express = require('express')
const session = require('express-session')
const msal = require('@azure/msal-node')

// MSAL App Configuration
const cryptoProvider = new msal.CryptoProvider()
const msalApp = new msal.ConfidentialClientApplication({
  auth: {
    authority: 'https://login.microsoftonline.com/consumers/',
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET
  }
})

const app = express()
app.use(express.json())
app.use(express.urlencoded({ extended: false }))
app.use(session({
  secret: process.env.COOKIE_SECRET,
  resave: false,
  saveUninitialized: false
}))

// Login Route
// This route will redirect the user to the Microsoft Identity Platform login page
// and then redirect back to the /redirect route after the user has logged in.
app.get('/login', async (req, res) => {
  const redirectUrl = await msalApp.getAuthCodeUrl({
    responseMode: 'form_post',
    redirectUri: 'http://localhost:8080/redirect',
    scopes: ['user.read'],
    // Encode the original URL the user was trying to access so we can redirect them back to it after logging in.
    state: cryptoProvider.base64Encode(JSON.stringify({originalUrl: req.query.originalUrl}))
  })

  res.redirect(redirectUrl)
})

// Redirect Route
// This route is called by the Microsoft Identity Platform after the user has logged in
// It will exchange the authorization code for an access token and ID token and then
// store them in the session before redirecting the user back to the original URL.
app.post('/redirect', async (req, res) => {
  const authResponse = await msalApp.acquireTokenByCode({
    code: req.body.code,
    scopes: ['user.read'],
    redirectUri: 'http://localhost:8080/redirect'
  }, req.body)

  req.session.authenticated = true
  req.session.account = authResponse.account
  req.session.idToken = authResponse.idToken
  req.session.user = authResponse.account.username

  // Check if we need to redirect back to the original url the user was trying to access.
  const successRedirectUrl = JSON.parse(atob(req.body.state)).originalUrl
  res.redirect(successRedirectUrl || '/')
})

// Logout route
app.get('/logout', (req, res) => {
  req.session.destroy()
  res.redirect('/')
})

// Auth Middleware - requires a user to be logged in to access any route after this point
app.use((req, res, next) => {
  const originalUrl = req.originalUrl

  if (!req.session || !req.session.authenticated) {
    return res.redirect(`/login?originalUrl=${originalUrl}`)
  }

  next()
})

// Catch all route just for demo purposes
app.get('*', (req, res) => {
  res.send(`Hello ${req.session.account.username}!`)
})

app.listen(8080, async () => {
  console.log('Server is running on http://localhost:8080')
})