# Contributing to Google Apps Scripts

Thank you for your interest in contributing! This document provides guidelines for contributing to this project.

## How to Contribute

### Reporting Bugs

1. Check if the bug has already been reported in [Issues](https://github.com/jjshay/google-apps-scripts/issues)
2. If not, create a new issue with:
   - Clear, descriptive title
   - Steps to reproduce
   - Expected vs actual behavior
   - Screenshots if applicable
   - Which script you're using

### Suggesting Features

1. Check existing issues for similar suggestions
2. Create a new issue with the "feature request" label
3. Describe the feature and its use case
4. Explain why it would be valuable

### Pull Requests

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Test in Google Apps Script editor
5. Commit with clear messages (`git commit -m 'Add amazing feature'`)
6. Push to your branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

## Development Setup

### Local Development

```bash
# Clone your fork
git clone https://github.com/YOUR_USERNAME/google-apps-scripts.git
cd google-apps-scripts

# Install clasp (Google Apps Script CLI)
npm install -g @google/clasp

# Login to Google
clasp login
```

### Testing Scripts

1. Create a test Google Sheet
2. Open Extensions > Apps Script
3. Copy/paste the script
4. Run and verify functionality

## Code Style

- Use meaningful variable and function names
- Add JSDoc comments to functions
- Keep functions focused and small
- Use `const` and `let` instead of `var`
- Handle errors gracefully with try/catch

## Script Documentation

When adding new scripts:
- Add a header comment describing the purpose
- Document required setup (API keys, sheet structure)
- Include example usage
- Update README.md with the new script

## Questions?

Feel free to open an issue for any questions or discussions!
