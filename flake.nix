{
  description = "Volunteerist - Google Apps Script volunteer follow-up tracker";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";
    flake-utils.url = "github:numtide/flake-utils";
  };

  outputs = { self, nixpkgs, flake-utils }:
    flake-utils.lib.eachDefaultSystem (system:
      let
        pkgs = nixpkgs.legacyPackages.${system};
      in
      {
        devShells.default = pkgs.mkShell {
          buildInputs = with pkgs; [
            biome
            nodejs_22
          ];

          shellHook = ''
            echo "Volunteerist dev environment"
            echo "Node.js $(node --version)"
            echo ""
            echo "Commands:"
            echo "  npm run build   - Compile TypeScript to dist/"
            echo "  npm run push    - Build and push to Apps Script"
            echo "  npm run test    - Run Vitest unit tests (also the pre-push hook)"
            echo "  npm run check:commit - Lint and typecheck (also the pre-commit hook)"
            echo "  npm run check   - Lint, typecheck, and test"
            echo "  npm run login   - Authenticate clasp with Google (interactive, once)"
            echo "  npm run login:status - Check clasp auth"
          '';
        };
      }
    );
}
