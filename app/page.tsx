import WorkOrganizer from "./work-organizer";
import { chatGPTSignInPath, getChatGPTUser } from "./chatgpt-auth";

export const dynamic = "force-dynamic";

export default async function Home() {
  const signedInUser = await getChatGPTUser();
  const user = signedInUser ?? (process.env.NODE_ENV !== "production"
    ? { displayName: "Nuno Moreira", email: "local-dev@work-organizer", fullName: "Nuno Moreira" }
    : null);

  if (!user) {
    return <main className="signin-page"><section className="signin-card"><span className="brand-mark">W</span><small>WORK ORGANIZER</small><h1>O teu planeamento fica privado.</h1><p>Inicia sessão para aceder às US, alocações e definições guardadas na tua base de dados.</p><a href={chatGPTSignInPath("/")}>Entrar com ChatGPT</a></section></main>;
  }

  return <WorkOrganizer user={{ displayName: user.displayName, email: user.email }} />;
}
