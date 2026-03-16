# Siemens Teamcenter / Bluebeam Studio PoC

## Disclaimer

This repository contains a proof-of-concept integration demonstrating how a system such as Teamcenter could interact with the Bluebeam Studio API.

This project is not an official Bluebeam product and is not supported by Bluebeam. It is provided solely as a reference implementation for evaluation and development purposes.

## What this PoC demonstrates

- OAuth authentication with the Bluebeam API
- Creating a Studio Session
- Uploading documents using the Bluebeam upload flow
- Inviting users to a Session
- Snapshot / roundtrip workflow support
- Cleanup of Session resources

## Architecture notes

The Bluebeam side of the workflow uses working API calls.
The Teamcenter side is stubbed/mocked to illustrate where ERP/PLM actions would occur.

## Requirements

- Node.js 18+ recommended
- Bluebeam API credentials
- A configured `.env` file based on `.env.example`

## Quick Start

1. Clone the repository
2. Copy `.env.example` to `.env`
3. Install dependencies

```bash
npm install
