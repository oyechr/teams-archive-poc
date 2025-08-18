import { useState, useEffect } from "react";
import { app } from "@microsoft/teams-js";

export function useTeamsContext() {
  const [inTeams, setInTeams] = useState<boolean>(false);
  const [context, setContext] = useState<any>(null);
  const [theme, setTheme] = useState<string>("default");

  useEffect(() => {
    app
      .initialize()
      .then(() => {
        setInTeams(true);
        app.getContext().then((ctx) => {
          setContext(ctx);
          setTheme(ctx.app.theme || "default");
        });
        app.registerOnThemeChangeHandler((theme) => {
          setTheme(theme || "default");
        });
      })
      .catch(() => {
        setInTeams(false);
      });
  }, []);

  return { inTeams, context, theme };
}
