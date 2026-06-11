properties([
	gitLabConnection('gitlab'),
	[$class: 'ParametersDefinitionProperty', 
		parameterDefinitions: [
			[$class: 'StringParameterDefinition', name: 'branch', defaultValue: 'master', description: 'the branch to build'],
			[$class: 'StringParameterDefinition', name: 'apiUrl', defaultValue: 'https://api-qa.aspose.cloud', description: 'api url'],
            [$class: 'BooleanParameterDefinition', name: 'ignoreCiSkip', defaultValue: false, description: 'ignore CI Skip'],
            [$class: 'StringParameterDefinition', name: 'credentialsId', defaultValue: '6839cbe8-39fa-40c0-86ce-90706f0bae5d', description: 'credentials id'],
            [$class: 'BooleanParameterDefinition', name: 'packageTesting', defaultValue: false, description: 'Testing package from repository without local sources. Used for prodhealthcheck'],
		]
	]
])

def needToBuild = false
def packageTesting = false

def runtests(dockerImageVersion)
{
    dir(dockerImageVersion){
        try {
				stage('checkout'){
					checkout([$class: 'GitSCM', branches: [[name: params.branch]], doGenerateSubmoduleConfigurations: false, extensions: [], submoduleCfg: [], userRemoteConfigs: [[credentialsId: '361885ba-9425-4230-950e-0af201d90547', url: 'https://git.auckland.dynabic.com/words-cloud/words-cloud-node.git']]])
					
                    sh 'git show -s HEAD > gitMessage'
                    def commitMessage = readFile('gitMessage').trim()
                    echo commitMessage
                    needToBuild = params.ignoreCiSkip || !commitMessage.contains('[ci skip]')
                    packageTesting = params.packageTesting                    
                    sh 'git clean -fdx'
                    
                    def apiUrl = params.apiUrl
                    
                    if (needToBuild) {
                        withCredentials([usernamePassword(credentialsId: params.credentialsId, passwordVariable: 'ClientSecret', usernameVariable: 'ClientId')]) {
                            sh 'echo "{\\"ClientId\\": \\"$ClientId\\",\\"ClientSecret\\": \\"$ClientSecret\\", \\"BaseUrl\\":\\"$apiUrl\\"}" > testConfig.json'
                        }
                    }
				}
            
            if (needToBuild) {
                docker.image('node:' + dockerImageVersion).inside{
                    if (packageTesting) {
                            stage('remove sources and redefine referencies'){
                                sh "npm uninstall asposewordscloud"
                                sh "sed -i 's/asposewordscloud/asposewordscloudtest/g' package.json"
                                sh "rm -rf src"
                                sh "find test -type f -name \"*.ts\" -exec sed -i 's+\".*/src/.*\"+\"asposewordscloud\"+g' {} +"
                                sh "find bdd -type f -name \"*.ts\" -exec sed -i 's+\".*/src/.*\"+\"asposewordscloud\"+g' {} +"
                                sh "npm install asposewordscloud" 
                            }
                    }
                
                        stage('build'){
                            withEnv([
                            /* Override the npm cache directory to avoid: EACCES: permission denied, mkdir '/.npm' */
                            'npm_config_cache=npm-cache',
                            /* set home to our current directory because other bower
                             * nonsense breaks with HOME=/, e.g.:
                             * EACCES: permission denied, mkdir '/.config'
                             */
                            'HOME=.',
                            ]) {
                                sh "npm install"
                                                        
                                sh "npm run tsc"
                                if (params.branch == 'refs/heads/master'){
                                    sh "npm run lint"
                                }
                            }
                        }
                    
                        stage('tests'){   
                            try {
                                sh "npm run test-jenkins"
                            } finally {
                                junit 'reports/**.xml'
                            }
                        }
                    
                        stage('bdd-tests'){
                            try {
                                sh "npm run cucumber"
                            } finally {
                                cucumber 'reports/**.json'
                            }
                        }
                } 
            }
        } finally {                       
            deleteDir()
			cleanWs()
			sh 'docker system prune -f'
        }
    }
}

node('words-linux') {        
    runtests("latest")   
}