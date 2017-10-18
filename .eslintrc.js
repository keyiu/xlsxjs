module.exports = {
  "parserOptions": {
    "ecmaVersion": 6,
  },
  "rules": {
    "no-undef": 2,
    "no-invalid-this": 0,
    "curly": 2,
    "max-len": [2, 120, 4],
  },
  "globals": {
    "process": true,
    "require": true,
    "module": true,
    "__dirname": true,
    "Promise": true,
    Map: true,
    describe: true,
    it: true,
  },
  "extends": "google"
};
