-- Optional IdP issuer validation for SAML SSO.
--
-- node-saml validates signatures/audience/time conditions but does not enforce
-- the IdP EntityID unless the application checks it. Store the expected IdP
-- issuer per provider so the API can reject assertions from an unexpected IdP.

ALTER TABLE org_saml_providers
  ADD COLUMN idp_issuer text;
