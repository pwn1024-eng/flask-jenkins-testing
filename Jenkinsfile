pipeline {
    agent any

    environment {
        SCRIPT  = 'newScrapper.py'
        APP_DIR = '/opt/scrapper'
        SERVICE = 'scrapper'
        PLAYWRIGHT_BROWSERS_PATH = '/opt/scrapper/browsers'
        SCAN_VENV = '/var/lib/jenkins/scan-venv'
    }

    options {
        timestamps()
        disableConcurrentBuilds()
        buildDiscarder(logRotator(numToKeepStr: '20'))
    }

    stages {
        stage('Syntax check') {
            steps { sh 'python3 -m py_compile ${SCRIPT}' }
        }

        stage('Security scan') {
            steps {
                sh '''
                    set -e
                    # scanner tools live in their own venv, installed once and kept current
                    [ -d ${SCAN_VENV} ] || python3 -m venv ${SCAN_VENV}
                    ${SCAN_VENV}/bin/pip install -q -U bandit pip-audit detect-secrets

                    echo "=== Bandit: code analysis ==="
                    ${SCAN_VENV}/bin/bandit -q -r ${SCRIPT} -f json -o bandit-report.json || true
                    ${SCAN_VENV}/bin/bandit -q -r ${SCRIPT}          # printed in the log
                    ${SCAN_VENV}/bin/bandit -q -r ${SCRIPT} -lll     # fails on HIGH severity only

                    echo "=== pip-audit: vulnerable dependencies ==="
                    ${SCAN_VENV}/bin/pip-audit -r requirements.txt --no-deps

                    echo "=== detect-secrets: committed secrets ==="
                    ${SCAN_VENV}/bin/detect-secrets scan --all-files \
                      --exclude-files '(\\.xlsx$|\\.git/)' > secrets-report.json
                    ${SCAN_VENV}/bin/python -c "
import json, sys
found = json.load(open('secrets-report.json'))['results']
for f, items in found.items():
    for i in items:
        print('SECRET?', f, 'line', i['line_number'], i['type'])
sys.exit(1 if found else 0)
"
                '''
            }
            post {
                always { archiveArtifacts artifacts: 'bandit-report.json, secrets-report.json', allowEmptyArchive: true }
            }
        }

        stage('Deploy') {
            when { branch 'main' }
            steps {
                sh '''
                    set -e
                    rsync -a ${SCRIPT} requirements.txt ${APP_DIR}/

                    if [ -d input ]; then
                        rsync -a --ignore-existing input/ ${APP_DIR}/input/
                    fi

                    cd ${APP_DIR}
                    [ -d venv ] || python3 -m venv venv
                    venv/bin/pip install -q -r requirements.txt
                    venv/bin/playwright install chromium
                    sudo systemctl restart ${SERVICE}
                '''
            }
        }

        stage('Verify') {
            when { branch 'main' }
            steps {
                sh '''
                    sleep 5
                    sudo systemctl is-active ${SERVICE}
                    journalctl -u ${SERVICE} -n 15 --no-pager || true
                '''
            }
        }
    }

    post {
        success { echo 'Deployed successfully' }
        failure { echo 'Pipeline failed, check the stage logs' }
    }
}