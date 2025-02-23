"use server";

export const sendDiscordMessage = async (message: string) => {
  try {
    // A fetch request to send data through the discord
    // webhook, and display it as a message in your
    // discord channel
    // await fetch(process.env.DISCORD_WEBHOOK_URL, {
    //   method: "POST",
    //   headers: {
    //     "Content-Type": "application/json",
    //   },
    //   body: JSON.stringify({
    //     content: message,
    //   }),
    // });
    console.log('ovde test');
    
  } catch (err: any) {
    // Just in case :)
    console.log(err.message);
  }
};
