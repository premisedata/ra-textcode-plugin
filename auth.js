/* global Office */

// Office 365 SSO Authentication Module

let currentUser = null;

async function getAccessToken() {
    try {
        // Use Office SSO to get token
        const token = await Office.auth.getAccessToken({ 
            allowSignInPrompt: true,
            allowConsentPrompt: true,
            forMSGraphAccess: true 
        });
        
        return token;
    } catch (error) {
        console.error('Error getting access token:', error);
        
        // Handle specific error codes
        if (error.code === 13001) {
            // User is not signed in
            throw new Error('Please sign in to Office 365');
        } else if (error.code === 13002) {
            // User consent required
            throw new Error('Please grant consent to use this add-in');
        } else if (error.code === 13003) {
            // SSO not supported
            throw new Error('Single Sign-On is not supported in this version of Office');
        }
        
        throw error;
    }
}

async function getUserProfile() {
    try {
        const token = await getAccessToken();
        
        // Call Microsoft Graph to get user info
        const response = await fetch('https://graph.microsoft.com/v1.0/me', {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        
        if (!response.ok) {
            throw new Error('Failed to fetch user profile');
        }
        
        const profile = await response.json();
        
        currentUser = {
            id: profile.id,
            email: profile.mail || profile.userPrincipalName,
            displayName: profile.displayName,
            givenName: profile.givenName,
            surname: profile.surname
        };
        
        return currentUser;
        
    } catch (error) {
        console.error('Error getting user profile:', error);
        throw error;
    }
}

function getCurrentUser() {
    return currentUser;
}

// Export functions
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { getAccessToken, getUserProfile, getCurrentUser };
}
