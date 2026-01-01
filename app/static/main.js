async function createUser() {
    const name = document.getElementById("name").value;
    const email = document.getElementById("email").value;

    const response = await fetch("/api/users/", {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            name: name,
            email: email
        })
    });

    const data = await response.json();
    document.getElementById("result").innerText =
        JSON.stringify(data, null, 2);
}