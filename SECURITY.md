# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| latest  | :white_check_mark: |

## Security Measures

This project implements the following security practices:

### CI/CD Pipeline
- **SAST** — Static Application Security Testing via [Bandit](https://bandit.readthedocs.io/)
- **Secret Scanning** — Automated secret detection via [Gitleaks](https://gitleaks.io/)
- **Dependency Audit** — Known vulnerability detection via [pip-audit](https://pypi.org/project/pip-audit/)
- **Container Scanning** — Docker image vulnerability scanning via [Trivy](https://trivy.dev/)
- **SBOM** — Software Bill of Materials generation via [Syft](https://github.com/anchore/syft)

### Development
- **Pre-commit hooks** — Automated checks before every commit (linting, secrets, security)
- **Pinned dependencies** — All dependencies specify minimum versions
- **Non-root Docker** — Container runs as unprivileged user

## Reporting a Vulnerability

If you discover a security vulnerability, please report it responsibly:

1. **Do NOT** open a public issue
2. Email: [your-email@example.com](mailto:your-email@example.com)
3. Include:
   - Description of the vulnerability
   - Steps to reproduce
   - Potential impact
4. Expected response time: **48 hours**

## Scope

This is a desktop utility for internal use. The primary threat model covers:
- **Supply chain** — Compromised dependencies
- **Data leakage** — Accidental commit of sensitive Excel/Word data
- **Build integrity** — Reproducible, verified builds
