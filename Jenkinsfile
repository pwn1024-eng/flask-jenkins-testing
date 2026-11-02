pipeline {
    agent any

    environment {
        DEPLOY_HOST = '192.168.0.12'
        DEPLOY_PORT = '22'                  // use 2222 if you used the netsh port proxy
        DEPLOY_USER = 'deploy'
        APP_DIR     = '/opt/scrapper'
        SERVICE     = 'scrapper'
    }

    triggers { githubPush() }               // fallback: pollSCM('H/2 * * * *')

    options {
        timestamps()
        disableConcurrentBuilds()
        buildDiscarder(logRotator(numToKeepStr: '20'))
    }

    stages {
        stage('Checkout') {
            steps { checkout scm }
        }

        stage('Syntax check') {
            steps { sh 'python3 -m py_compile newScrapper.py' }
        }

        stage('Deploy') {
            when { branch 'main' }
            steps {
                sshagent(credentials: ['wsl-deploy-key']) {
                    sh '''
                        SSH="ssh -p ${DEPLOY_PORT} -o StrictHostKeyChecking=accept-new ${DEPLOY_USER}@${DEPLOY_HOST}"
                        RSYNC_SSH="ssh -p ${DEPLOY_PORT} -o StrictHostKeyChecking=accept-new"

                        # code only; never touch live data or the venv
                        rsync -az -e "$RSYNC_SSH" \
                          --exclude '.git' --exclude 'venv' \
                          --exclude 'input/' --exclude 'output/' \
                          newScrapper.py requirements.txt \
                          ${DEPLOY_USER}@${DEPLOY_HOST}:${APP_DIR}/

                        # seed input files only if they don't exist yet (the app edits them in place)
                        [ -d input ] && rsync -az --ignore-existing -e "$RSYNC_SSH" input/ \
                          ${DEPLOY_USER}@${DEPLOY_HOST}:${APP_DIR}/input/ || true

                        $SSH "cd ${APP_DIR} && \
                              [ -d venv ] || python3 -m venv venv && \
                              venv/bin/pip install -q -r requirements.txt && \
                              venv/bin/playwright install chromium && \
                              sudo systemctl restart ${SERVICE}"
                    '''
                }
            }
        }

        stage('Verify') {
            when { branch 'main' }
            steps {
                sshagent(credentials: ['wsl-deploy-key']) {
                    sh '''
                        sleep 5
                        ssh -p ${DEPLOY_PORT} ${DEPLOY_USER}@${DEPLOY_HOST} \
                          "systemctl is-active ${SERVICE} && journalctl -u ${SERVICE} -n 15 --no-pager"
                    '''
                }
            }
        }
    }

    post {
        failure { echo 'Deploy failed, check the stage logs' }
    }
}