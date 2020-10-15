const { argv, env } = require('process');

var exec = require('child_process').exec;

var action = null;
var killTeamsCommand = "taskkill /f /im teams.exe";

var shouldDeployTeams = false;

// check that we have valid params
if (argv && argv.length >= 3) {
  action = argv[2];
  shouldDeployTeams = action == "deploy" && argv.length >= 4;
} else {
  console.error("Usage: {action} {deploy command}?");
}

exec(killTeamsCommand, function(err, stdout, stderr) {
  console.log("Closing Teams");

  if (shouldDeployTeams) {
    var localAppData = env.localAppData;
    var workspaceRoot = argv[3];

    var deployTeamsCommand = localAppData + '/Microsoft/Teams/Update.exe --processStart "teams.exe" --process-start-args "--installAppPackage=""' + workspaceRoot + '/.publish/Development.zip"""';
    exec(deployTeamsCommand, (e, out, err) => {
      if (out) console.log(out);
      if (err) console.error(err);
      console.log("Launching Teams to deploy app package.");
    });
  }
});