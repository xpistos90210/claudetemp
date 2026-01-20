// Game Constants
const FPS = 60;
const SHIP_SIZE = 20;
const SHIP_THRUST = 5;
const SHIP_TURN_SPEED = 360;
const FRICTION = 0.7;
const BULLET_SPEED = 500;
const BULLET_MAX = 10;
const BULLET_LIFE = 1;
const ASTEROID_SPEED = 50;
const ASTEROID_VERTICES = 10;
const ASTEROID_JAG = 0.4;
const ASTEROID_POINTS_LARGE = 20;
const ASTEROID_POINTS_MEDIUM = 50;
const ASTEROID_POINTS_SMALL = 100;
const STARTING_LIVES = 3;
const INVINCIBILITY_TIME = 3;
const PARTICLES_PER_EXPLOSION = 20;

// Game State
let canvas, ctx;
let ship, asteroids, bullets, particles;
let score, lives, level, highScore;
let gameRunning = false;
let invincibilityTimer = 0;

// Controls
let keys = {
    left: false,
    right: false,
    thrust: false,
    fire: false
};
let canFire = true;

// Initialize
function init() {
    canvas = document.getElementById('gameCanvas');
    ctx = canvas.getContext('2d');

    resizeCanvas();
    window.addEventListener('resize', resizeCanvas);

    // Load high score
    highScore = localStorage.getItem('asteroids-highscore') || 0;
    document.getElementById('high-score').textContent = highScore;

    setupControls();

    // Start button
    document.getElementById('start-btn').addEventListener('click', startGame);
    document.getElementById('restart-btn').addEventListener('click', startGame);

    // Initial render
    requestAnimationFrame(gameLoop);
}

function resizeCanvas() {
    canvas.width = window.innerWidth * window.devicePixelRatio;
    canvas.height = window.innerHeight * window.devicePixelRatio;
    ctx.scale(window.devicePixelRatio, window.devicePixelRatio);
}

function setupControls() {
    // Keyboard controls
    document.addEventListener('keydown', (e) => {
        switch(e.key) {
            case 'ArrowLeft':
            case 'a':
                keys.left = true;
                break;
            case 'ArrowRight':
            case 'd':
                keys.right = true;
                break;
            case 'ArrowUp':
            case 'w':
                keys.thrust = true;
                break;
            case ' ':
                keys.fire = true;
                break;
        }
    });

    document.addEventListener('keyup', (e) => {
        switch(e.key) {
            case 'ArrowLeft':
            case 'a':
                keys.left = false;
                break;
            case 'ArrowRight':
            case 'd':
                keys.right = false;
                break;
            case 'ArrowUp':
            case 'w':
                keys.thrust = false;
                break;
            case ' ':
                keys.fire = false;
                canFire = true;
                break;
        }
    });

    // Touch controls
    const touchHandler = (id, key, isTouch) => {
        const el = document.getElementById(id);

        const start = (e) => {
            e.preventDefault();
            keys[key] = true;
            el.classList.add('active');
            if (navigator.vibrate) navigator.vibrate(10);
        };

        const end = (e) => {
            e.preventDefault();
            keys[key] = false;
            el.classList.remove('active');
            if (key === 'fire') canFire = true;
        };

        el.addEventListener('touchstart', start, { passive: false });
        el.addEventListener('touchend', end, { passive: false });
        el.addEventListener('touchcancel', end, { passive: false });
        el.addEventListener('mousedown', start);
        el.addEventListener('mouseup', end);
        el.addEventListener('mouseleave', end);
    };

    touchHandler('rotate-left', 'left');
    touchHandler('rotate-right', 'right');
    touchHandler('thrust', 'thrust');
    touchHandler('fire', 'fire');
}

function startGame() {
    document.getElementById('start-screen').classList.add('hidden');
    document.getElementById('game-over').classList.add('hidden');

    score = 0;
    lives = STARTING_LIVES;
    level = 1;
    gameRunning = true;

    updateScore();
    updateLives();

    createShip();
    asteroids = [];
    bullets = [];
    particles = [];

    createAsteroids();
}

function createShip() {
    const w = window.innerWidth;
    const h = window.innerHeight;

    ship = {
        x: w / 2,
        y: h / 2,
        radius: SHIP_SIZE / 2,
        angle: -Math.PI / 2,
        rotation: 0,
        thrusting: false,
        thrust: { x: 0, y: 0 },
        canShoot: true
    };

    invincibilityTimer = INVINCIBILITY_TIME;
}

function createAsteroids() {
    const w = window.innerWidth;
    const h = window.innerHeight;
    const numAsteroids = level + 3;

    for (let i = 0; i < numAsteroids; i++) {
        let x, y;
        // Don't spawn too close to ship
        do {
            x = Math.random() * w;
            y = Math.random() * h;
        } while (distBetweenPoints(ship.x, ship.y, x, y) < 150);

        asteroids.push(createAsteroid(x, y, Math.ceil(SHIP_SIZE * 3)));
    }
}

function createAsteroid(x, y, radius) {
    const lvlMult = 1 + 0.1 * level;
    const speed = ASTEROID_SPEED * lvlMult;

    return {
        x: x,
        y: y,
        xv: (Math.random() * speed - speed / 2) / FPS,
        yv: (Math.random() * speed - speed / 2) / FPS,
        radius: radius,
        angle: Math.random() * Math.PI * 2,
        vertices: Math.floor(Math.random() * (ASTEROID_VERTICES + 1) + ASTEROID_VERTICES / 2),
        offsets: Array(ASTEROID_VERTICES + 5).fill(0).map(() => Math.random() * ASTEROID_JAG * 2 + 1 - ASTEROID_JAG)
    };
}

function shootBullet() {
    if (!canFire || bullets.length >= BULLET_MAX) return;

    bullets.push({
        x: ship.x + Math.cos(ship.angle) * SHIP_SIZE / 2,
        y: ship.y + Math.sin(ship.angle) * SHIP_SIZE / 2,
        xv: BULLET_SPEED * Math.cos(ship.angle) / FPS,
        yv: BULLET_SPEED * Math.sin(ship.angle) / FPS,
        life: BULLET_LIFE
    });

    canFire = false;
    if (navigator.vibrate) navigator.vibrate(5);
}

function createExplosion(x, y, color = '#0ff') {
    for (let i = 0; i < PARTICLES_PER_EXPLOSION; i++) {
        const angle = Math.random() * Math.PI * 2;
        const speed = Math.random() * 3 + 1;
        particles.push({
            x: x,
            y: y,
            xv: Math.cos(angle) * speed,
            yv: Math.sin(angle) * speed,
            life: 1,
            color: color,
            size: Math.random() * 3 + 1
        });
    }
}

function destroyAsteroid(index) {
    const asteroid = asteroids[index];

    // Score
    if (asteroid.radius > SHIP_SIZE * 1.5) {
        score += ASTEROID_POINTS_LARGE;
    } else if (asteroid.radius > SHIP_SIZE * 0.75) {
        score += ASTEROID_POINTS_MEDIUM;
    } else {
        score += ASTEROID_POINTS_SMALL;
    }
    updateScore();

    // Explosion
    createExplosion(asteroid.x, asteroid.y);
    if (navigator.vibrate) navigator.vibrate(20);

    // Split into smaller asteroids
    if (asteroid.radius > SHIP_SIZE) {
        asteroids.push(createAsteroid(asteroid.x, asteroid.y, asteroid.radius / 2));
        asteroids.push(createAsteroid(asteroid.x, asteroid.y, asteroid.radius / 2));
    }

    asteroids.splice(index, 1);

    // Next level
    if (asteroids.length === 0) {
        level++;
        createAsteroids();
    }
}

function shipHit() {
    if (invincibilityTimer > 0) return;

    createExplosion(ship.x, ship.y, '#f66');
    if (navigator.vibrate) navigator.vibrate([50, 50, 50]);

    lives--;
    updateLives();

    if (lives <= 0) {
        gameOver();
    } else {
        createShip();
    }
}

function gameOver() {
    gameRunning = false;

    if (score > highScore) {
        highScore = score;
        localStorage.setItem('asteroids-highscore', highScore);
    }

    document.getElementById('final-score').textContent = score;
    document.getElementById('high-score').textContent = highScore;
    document.getElementById('game-over').classList.remove('hidden');
}

function updateScore() {
    document.getElementById('score').textContent = score;
}

function updateLives() {
    const livesEl = document.getElementById('lives');
    livesEl.innerHTML = '';
    for (let i = 0; i < lives; i++) {
        const life = document.createElement('div');
        life.className = 'life-icon';
        livesEl.appendChild(life);
    }
}

function distBetweenPoints(x1, y1, x2, y2) {
    return Math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2);
}

function wrapPosition(obj) {
    const w = window.innerWidth;
    const h = window.innerHeight;

    if (obj.x < 0 - obj.radius) obj.x = w + obj.radius;
    else if (obj.x > w + obj.radius) obj.x = 0 - obj.radius;

    if (obj.y < 0 - obj.radius) obj.y = h + obj.radius;
    else if (obj.y > h + obj.radius) obj.y = 0 - obj.radius;
}

function update(deltaTime) {
    if (!gameRunning) return;

    const dt = deltaTime / 1000;

    // Update invincibility
    if (invincibilityTimer > 0) {
        invincibilityTimer -= dt;
    }

    // Ship rotation
    if (keys.left) {
        ship.angle -= SHIP_TURN_SPEED * dt * Math.PI / 180;
    }
    if (keys.right) {
        ship.angle += SHIP_TURN_SPEED * dt * Math.PI / 180;
    }

    // Ship thrust
    ship.thrusting = keys.thrust;
    if (ship.thrusting) {
        ship.thrust.x += SHIP_THRUST * Math.cos(ship.angle) * dt;
        ship.thrust.y += SHIP_THRUST * Math.sin(ship.angle) * dt;
    } else {
        ship.thrust.x *= Math.pow(FRICTION, dt * 10);
        ship.thrust.y *= Math.pow(FRICTION, dt * 10);
    }

    // Move ship
    ship.x += ship.thrust.x;
    ship.y += ship.thrust.y;
    wrapPosition(ship);

    // Fire
    if (keys.fire) {
        shootBullet();
    }

    // Update bullets
    for (let i = bullets.length - 1; i >= 0; i--) {
        bullets[i].x += bullets[i].xv;
        bullets[i].y += bullets[i].yv;
        bullets[i].life -= dt;

        // Wrap bullets
        wrapPosition({ ...bullets[i], radius: 0 });

        if (bullets[i].life <= 0) {
            bullets.splice(i, 1);
        }
    }

    // Update asteroids
    for (let asteroid of asteroids) {
        asteroid.x += asteroid.xv;
        asteroid.y += asteroid.yv;
        wrapPosition(asteroid);
    }

    // Update particles
    for (let i = particles.length - 1; i >= 0; i--) {
        particles[i].x += particles[i].xv;
        particles[i].y += particles[i].yv;
        particles[i].life -= dt * 2;

        if (particles[i].life <= 0) {
            particles.splice(i, 1);
        }
    }

    // Collision detection - bullets vs asteroids
    for (let i = bullets.length - 1; i >= 0; i--) {
        for (let j = asteroids.length - 1; j >= 0; j--) {
            if (distBetweenPoints(bullets[i].x, bullets[i].y, asteroids[j].x, asteroids[j].y) < asteroids[j].radius) {
                bullets.splice(i, 1);
                destroyAsteroid(j);
                break;
            }
        }
    }

    // Collision detection - ship vs asteroids
    if (invincibilityTimer <= 0) {
        for (let asteroid of asteroids) {
            if (distBetweenPoints(ship.x, ship.y, asteroid.x, asteroid.y) < ship.radius + asteroid.radius * 0.8) {
                shipHit();
                break;
            }
        }
    }
}

function draw() {
    const w = window.innerWidth;
    const h = window.innerHeight;

    // Clear
    ctx.fillStyle = 'transparent';
    ctx.clearRect(0, 0, w, h);

    if (!gameRunning && !document.getElementById('start-screen').classList.contains('hidden')) {
        return;
    }

    // Draw particles
    for (let particle of particles) {
        ctx.globalAlpha = particle.life;
        ctx.fillStyle = particle.color;
        ctx.beginPath();
        ctx.arc(particle.x, particle.y, particle.size, 0, Math.PI * 2);
        ctx.fill();
    }
    ctx.globalAlpha = 1;

    // Draw asteroids
    ctx.strokeStyle = '#0ff';
    ctx.lineWidth = 2;
    for (let asteroid of asteroids) {
        ctx.beginPath();
        for (let i = 0; i < asteroid.vertices; i++) {
            const angle = asteroid.angle + (i * Math.PI * 2 / asteroid.vertices);
            const r = asteroid.radius * asteroid.offsets[i];
            const x = asteroid.x + r * Math.cos(angle);
            const y = asteroid.y + r * Math.sin(angle);

            if (i === 0) {
                ctx.moveTo(x, y);
            } else {
                ctx.lineTo(x, y);
            }
        }
        ctx.closePath();
        ctx.stroke();

        // Glow effect
        ctx.shadowColor = '#0ff';
        ctx.shadowBlur = 10;
        ctx.stroke();
        ctx.shadowBlur = 0;
    }

    // Draw bullets
    ctx.fillStyle = '#fff';
    ctx.shadowColor = '#fff';
    ctx.shadowBlur = 10;
    for (let bullet of bullets) {
        ctx.beginPath();
        ctx.arc(bullet.x, bullet.y, 3, 0, Math.PI * 2);
        ctx.fill();
    }
    ctx.shadowBlur = 0;

    // Draw ship
    if (gameRunning) {
        // Blinking when invincible
        if (invincibilityTimer > 0 && Math.floor(invincibilityTimer * 10) % 2 === 0) {
            ctx.globalAlpha = 0.3;
        }

        ctx.strokeStyle = '#0ff';
        ctx.lineWidth = 2;
        ctx.shadowColor = '#0ff';
        ctx.shadowBlur = 15;

        ctx.beginPath();
        // Nose
        ctx.moveTo(
            ship.x + SHIP_SIZE * Math.cos(ship.angle),
            ship.y + SHIP_SIZE * Math.sin(ship.angle)
        );
        // Left corner
        ctx.lineTo(
            ship.x + SHIP_SIZE * 0.6 * Math.cos(ship.angle + Math.PI * 0.8),
            ship.y + SHIP_SIZE * 0.6 * Math.sin(ship.angle + Math.PI * 0.8)
        );
        // Back center
        ctx.lineTo(
            ship.x + SHIP_SIZE * 0.3 * Math.cos(ship.angle + Math.PI),
            ship.y + SHIP_SIZE * 0.3 * Math.sin(ship.angle + Math.PI)
        );
        // Right corner
        ctx.lineTo(
            ship.x + SHIP_SIZE * 0.6 * Math.cos(ship.angle - Math.PI * 0.8),
            ship.y + SHIP_SIZE * 0.6 * Math.sin(ship.angle - Math.PI * 0.8)
        );
        ctx.closePath();
        ctx.stroke();

        // Draw thrust flame
        if (ship.thrusting) {
            ctx.strokeStyle = '#f80';
            ctx.shadowColor = '#f80';
            ctx.beginPath();
            ctx.moveTo(
                ship.x + SHIP_SIZE * 0.3 * Math.cos(ship.angle + Math.PI * 0.9),
                ship.y + SHIP_SIZE * 0.3 * Math.sin(ship.angle + Math.PI * 0.9)
            );
            ctx.lineTo(
                ship.x + SHIP_SIZE * (0.5 + Math.random() * 0.3) * Math.cos(ship.angle + Math.PI),
                ship.y + SHIP_SIZE * (0.5 + Math.random() * 0.3) * Math.sin(ship.angle + Math.PI)
            );
            ctx.lineTo(
                ship.x + SHIP_SIZE * 0.3 * Math.cos(ship.angle - Math.PI * 0.9),
                ship.y + SHIP_SIZE * 0.3 * Math.sin(ship.angle - Math.PI * 0.9)
            );
            ctx.stroke();
        }

        ctx.shadowBlur = 0;
        ctx.globalAlpha = 1;
    }
}

let lastTime = 0;
function gameLoop(timestamp) {
    const deltaTime = timestamp - lastTime;
    lastTime = timestamp;

    update(Math.min(deltaTime, 100));
    draw();

    requestAnimationFrame(gameLoop);
}

// Start when DOM is ready
document.addEventListener('DOMContentLoaded', init);
