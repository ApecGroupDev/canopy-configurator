#!/bin/bash
# Railway startup script for Canopy Configurator
# Writes Streamlit secrets from environment variable before launching the app.

set -e

# If GCP_SECRETS_TOML is set (Railway env var), write it to secrets.toml
# so st.secrets["gcp_service_account"] works as normal.
if [ -n "$GCP_SECRETS_TOML" ]; then
    mkdir -p .streamlit
    echo "$GCP_SECRETS_TOML" > .streamlit/secrets.toml
    echo "Secrets file written from GCP_SECRETS_TOML env var."
else
    echo "GCP_SECRETS_TOML not set — Google Sheets tracker will fall back to local CSV."
fi

# Start Streamlit
exec streamlit run canopy_configurator.py \
    --server.port "${PORT:-8501}" \
    --server.address 0.0.0.0 \
    --server.headless true
