# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 1.0.x   | :white_check_mark: |

## Reporting a Vulnerability

We take security seriously. If you discover a security vulnerability, please follow these steps:

### How to Report

1. **Do NOT** open a public GitHub issue for security vulnerabilities
2. Email the maintainer directly with details of the vulnerability
3. Include:
   - Description of the vulnerability
   - Steps to reproduce
   - Potential impact
   - Suggested fix (if any)

### What to Expect

- **Initial Response**: Within 48 hours
- **Status Update**: Within 7 days
- **Resolution Timeline**: Depends on severity
  - Critical: 24-48 hours
  - High: 7 days
  - Medium: 30 days
  - Low: 90 days

### Security Best Practices for Users

1. **Review script permissions**: Only grant necessary Google Workspace scopes
2. **Audit script access**: Regularly review which accounts have access
3. **Use project-specific deployments**: Don't share scripts across unrelated projects
4. **Monitor execution logs**: Watch for unauthorized or unexpected executions
5. **Protect sensitive data**: Be careful with data accessed via Google APIs

## Security Features

This project includes:

- **Dependabot**: Automated security updates for dependencies
- **CodeQL**: Static analysis for vulnerability detection

## Google Apps Script Security

### OAuth Scopes
- Request only the minimum required scopes
- Review scopes before deploying scripts
- Document all required scopes in README

### Data Handling
- Never log sensitive data (PII, passwords, tokens)
- Use PropertiesService for storing secrets
- Validate and sanitize all user inputs

### Deployment
- Use versioned deployments for production
- Test thoroughly in development before deploying
- Keep deployment URLs private when appropriate

### Access Control
- Limit script execution to authorized users
- Use Google Groups for managing access at scale
- Regularly audit who has access to scripts and data

## Acknowledgments

We appreciate responsible disclosure and will acknowledge security researchers who help improve this project's security.
